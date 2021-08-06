# SPDX-License-Identifier: MIT

"""
This module contains functions for making system data.
"""

import win32com.client

from typing import Sequence

import numpy as np
import pandas as pd

import monitor_channel
import monitor_factory

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def make_load_voltage_data(
    dss: win32com.client.CDispatch, monitors: Sequence[monitor_factory.Monitor]
) -> pd.DataFrame:
    """
    Make the load voltage data.

    Parameters
    ----------
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.
    monitors : list of monitor_factory.Monitor
        The load monitors.

    Returns
    -------
    pandas.core.frame.DataFrame, (n_timestep, n_load)
        The voltage magnitude data for the loads.
    """
    load_voltage_df = monitor_factory.make_monitor_data(
        channel=monitor_channel.Load.MODE_0_V1.value, dss=dss, monitors=monitors
    )

    return load_voltage_df
