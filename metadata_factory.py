# SPDX-License-Identifier: BSD-3-Clause

"""
This script contains functions for creating metadata for the EPRI ckt5 test
circuit.
"""

import win32com.client

from typing import List, Sequence, Tuple

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
def make_metadata(dss: win32com.client.CDispatch):
    """
    Make the metadata for the loads.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.

    Returns
    -------
    pandas.DataFrame, (n_loads, n_fields)
        The load metadata.
    """
    #***********************************************************************************************
    # Make the various metadata pieces.
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
    # Make the metadata data frame.
    #***********************************************************************************************
    metadata_df = pd.DataFrame(
        data={
            "load_name": load_names,
            "phase": load_phases,
            "loadshape": loadshapes,
            "transformer_name": load_transformer_names
        }
    )

    return metadata_df
