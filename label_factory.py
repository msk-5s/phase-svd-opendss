# SPDX-License-Identifier: BSD-3-Clause

"""
This script contains functions for creating labels for the EPRI ckt5 test circuit.
"""

import win32com.client

from typing import List, Sequence, Tuple

import numpy as np
import pandas as pd

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def _make_line_buses(dss: win32com.client.CDispatch) -> Tuple[List[str], List[str]]:
    """
    Make the bus names for all of the lines in the circuit.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.

    Returns
    -------
    line_bus1s : list of str, (n_lines,)
        The `bus1` of the lines.
    line_bus2s : list of str, (n_lines,)
        The `bus2` of the lines.
    """
    line_bus1s = []
    line_bus2s = []

    i_lines = dss.ActiveCircuit.Lines
    line_names = list(i_lines.AllNames)

    for name in line_names:
        i_lines.Name = name

        line_bus1s.append(str(i_lines.Bus1))
        line_bus2s.append(str(i_lines.Bus2))

    return (line_bus1s, line_bus2s)

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def _make_load_buses(dss: win32com.client.CDispatch, load_names: Sequence[str]) -> List[str]:
    """
    Make the name of load buses.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.
    load_names : list of str, (n_loads,)
        The load names.

    Returns
    -------
    list of str, (n_loads,)
        The load bus names.
    """
    load_buses = []

    for name in load_names:
        # The `ILoads` COM interface doesn't have a `Bus1` property.
        dss.Text.Command = f"? load.{name}.bus1"
        load_bus = str(dss.Text.Result)

        load_buses.append(load_bus)

    return load_buses

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def _make_loadshapes(dss: win32com.client.CDispatch, load_names: Sequence[str]) -> List[str]:
    """
    Make the names of the yearly loadshape for each load.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.
    load_names : list of str, (n_loads,)
        The load names.

    Returns
    -------
    list of str, (n_loads,)
        The loadshape for each load.
    """
    loadshapes = []

    for name in load_names:
        dss.ActiveCircuit.Loads.name = name

        loadshapes.append(str(dss.ActiveCircuit.Loads.Yearly))

    return loadshapes

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def _make_transformer_buses(dss: win32com.client.CDispatch) -> List[str]:
    """
    Make the bus names for the active winding of all the transformers in the
    circuit.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.

    Returns
    -------
    list of str, (n_transformers,)
        The bus names for the active winding of each transformer.
    """
    transformer_buses = []

    transformer_names = list(dss.ActiveCircuit.Transformers.AllNames)

    for name in transformer_names:
        # The `ITransformers` COM interface doesn't have a `Bus` property.
        dss.Text.Command = f"? transformer.{name}.bus"
        bus = str(dss.Text.Result)

        transformer_buses.append(bus)

    return transformer_buses

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def _make_transformer_values( # pylint: disable=too-many-locals
        dss: win32com.client.CDispatch,
        line_bus1s: Sequence[str],
        line_bus2s: Sequence[str],
        load_buses: Sequence[str],
        transformer_buses: Sequence[str]
    ) -> List[str]:
    """
    Make the names of the transformers connected to each load.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.
    line_bus1s : list of str, (n_lines,)
        The `bus1` of the lines.
    line_bus2s : list of str, (n_lines,)
        The `bus2` of the lines.
    load_buses : list of str, (n_loads,)
        The load bus names.
    transformer_buses : list of str, (n_transformers,)
        The bus names for the active winding of each transformer.

    Returns
    -------
    list of str, (n_load,)
        The name of the transformer each load is connected to.
    """
    load_transformer_names = []

    i_transformers = dss.ActiveCircuit.Transformers
    transformer_names = list(i_transformers.AllNames)

    for load_bus in load_buses:
        # Get the name of the line's `bus1` that the load is connected to.
        # `bus2` of a line is connected to the load and `bus1` of the line is
        # connected to the respective transformer.
        i = line_bus2s.index(load_bus)
        line_bus1 = line_bus1s[i]

        # Get the name of the transformer that `bus1` of the line is connected
        # to. This gives us the transformer that the load is connected to.
        i = transformer_buses.index(line_bus1)
        transformer_name = transformer_names[i]

        load_transformer_names.append(transformer_name)

    return load_transformer_names

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def make_load_labels(dss: win32com.client.CDispatch) -> pd.DataFrame:
    """
    Make the labels for the loads.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.

    Returns
    -------
    pandas.DataFrame, (n_loads, n_fields)
        The labels.
    """
    #***********************************************************************************************
    # Make the various label pieces.
    #***********************************************************************************************
    load_names = list(dss.ActiveCircuit.Loads.AllNames)
    load_buses = _make_load_buses(dss=dss, load_names=load_names)

    # Since these are Y-connected loads, there is only a single terminal
    # connected for the respective phase. Phase A = 1, B = 2, C = 3.
    # We subtract 1 so that the phases start at 0. Phase A = 0, etc.
    load_phases = [int(bus.split(sep=".")[1]) - 1 for bus in load_buses]

    loadshapes = _make_loadshapes(dss=dss, load_names=load_names)
    transformer_buses = _make_transformer_buses(dss=dss)

    (line_bus1s, line_bus2s) = _make_line_buses(dss=dss)

    load_transformer_names = _make_transformer_values(
        dss=dss,
        line_bus1s=line_bus1s,
        line_bus2s=line_bus2s,
        load_buses=load_buses,
        transformer_buses=transformer_buses
    )

    #***********************************************************************************************
    # Make the label data frame.
    #***********************************************************************************************
    label_df = pd.DataFrame(
        data={
            "load_name": load_names,
            "phase": load_phases,
            "loadshape": loadshapes,
            "transformer_name": load_transformer_names
        }
    )

    return label_df

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def make_xfmr_labels(dss: win32com.client.CDispatch, labels_load_df: pd.DataFrame) -> pd.DataFrame:
    """
    Make the labels for the xfmrs.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.

    Returns
    -------
    pandas.DataFrame, (n_xfmr, n_fields)
        The labels.
    """
    #***********************************************************************************************
    # Get the phase of the 591 secondary distribution transformers.
    #***********************************************************************************************
    # There are only 591 distribution transformers, but each transformer can have multiple loads.
    unique_xfmr_names = np.unique(labels_load_df["transformer_name"].to_numpy(dtype=str))
    load_xfmr_names = labels_load_df["transformer_name"].to_numpy(dtype=str)

    # Get the indices for the loads connected to each transformer.
    xfmr_load_indices_list = [
        np.where(load_xfmr_names == xfmr_name)[0] for xfmr_name in unique_xfmr_names
    ]

    # Since the transformers are single phase, we only need to get the phase of the first load to
    # get the phase of the transformer.
    phases = [
        labels_load_df.at[xfmr_load_indices[0], "phase"]
    for xfmr_load_indices in xfmr_load_indices_list]

    #***********************************************************************************************
    # Get the names and indices of the loads connected to each secondary distribution transformer.
    #***********************************************************************************************
    load_names = [
        ";".join(labels_load_df.loc[xfmr_load_indices, "load_name"].to_list())
    for xfmr_load_indices in xfmr_load_indices_list]

    load_indices = [
        ";".join([str(i) for i in xfmr_load_indices])
    for xfmr_load_indices in xfmr_load_indices_list]

    #***********************************************************************************************
    # Make the label data frame.
    #***********************************************************************************************
    labels_df = pd.DataFrame(
        data={
            "load_indices": load_indices,
            "load_names": load_names,
            "phase": phases,
            "transformer_name": unique_xfmr_names
        }
    )

    return labels_df
