# SPDX-License-Identifier: MIT

"""
This module contains functions for making monitors.
"""

from typing import List, NamedTuple, Sequence

import win32com.client

from rich.progress import track

import pandas as pd

#===================================================================================================
#===================================================================================================
class Monitor(NamedTuple):
    """
    A tuple of monitor information.

    Attributes
    ----------
    dss_command : str
        The OpenDSS command that will create the monitor.
    object_name : str
        The name of the OpenDSS object being monitored formated as classname.elementname.
    name : str
        The name of the monitor.
    """
    dss_command: str
    object_name: str
    name: str

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def make_monitors(object_names: Sequence[str], mode: int, terminal: int) -> List[Monitor]:
    """
    Make monitors for elements.

    Parameters
    ----------
    object_names : list of str
        The name of the OpenDSS objects to be monitored.
    mode : int
        The mode of the monitors (see the OpenDSS manual for the different modes).
    terminal : int
        The terminal of the object to monitor.

    Returns
    -------
    list of monitor_factory.Monitor
        The new monitors for each object.
    """
    monitors = []

    for object_name in track(object_names, "Making monitors..."):
        element_name = object_name.split(".")[-1]
        name = f"{element_name}_mode_{mode}_terminal_{terminal}"

        dss_command = f"new monitor.{name} element={object_name} mode={mode} terminal={terminal}"

        monitor = Monitor(dss_command=dss_command, object_name=object_name, name=name)

        monitors.append(monitor)

    return monitors

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def make_monitor_data(channel: int, dss: win32com.client.CDispatch, monitors: Sequence[Monitor])\
    -> pd.DataFrame:
    """
    Read the `monitors` and place their data in a data frame.

    Parameters
    ----------
    channel : int
        The monitor channel to read.
    dss : win32com.client.CDispatch
        The OpenDSSEngine COM object.
    monitors : list of monitor_factory.Monitor
        The monitors the read data from.

    Returns
    -------
    pandas.DataFrame, (n_timestep, n_monitor)
        A dataframe containing the data from each monitor for the given `channel`.
    """
    i_monitors = dss.ActiveCircuit.Monitors

    # We assume that all of the provided `monitors` record the same number of timesteps.
    i_monitors.Name = monitors[0].name
    hours = list(i_monitors.dblHour)
    element_names = [monitor.object_name.split(".")[-1] for monitor in monitors]

    frame = pd.DataFrame(index=hours, data=0.0, columns=element_names)

    for (monitor, element_name) in track(
        zip(monitors, element_names), "Making monitor data...", total=len(monitors)
    ):
        i_monitors.Name = monitor.name
        data = list(i_monitors.Channel(channel))

        frame[element_name] = data

    return frame
