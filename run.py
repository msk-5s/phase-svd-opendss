# SPDX-License-Identifier: MIT

"""
This script generates a synthetic dataset from EPRI's ckt5 test feeder circuit.
"""

import os
import win32com.client

import numpy as np
import pandas as pd
import pyarrow.feather

import data_factory
import metadata_factory
import monitor_factory
import profile_factory

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def main(): # pylint: disable=too-many-locals
    """
    The main function.
    """
    # OpenDSS prefers full paths. Use the current directory of this script file.
    basepath = os.getcwd()

    #***********************************************************************************************
    # Get the OpenDSS engine object and load the circuit.
    #***********************************************************************************************
    print("Loading circuit...")

    dss = win32com.client.Dispatch("OpenDSSEngine.DSS")

    dss.Text.Command = "clearall"
    dss.Text.Command = f"redirect ({basepath}/ckt5-src/Master_ckt5.dss)"

    #***********************************************************************************************
    # Make the metadata.
    #***********************************************************************************************
    # NOTE: It's important to make the metadata before assigning the synthetic load profiles to the
    # loads in the circuit. We want to have the names of the original load profiles before changing
    # them so that we will know what each type of load is (Residential, Commercial_SM,
    # Commercial_MD).
    print("Making metadata...")

    metadata_df = metadata_factory.make_metadata(dss=dss)

    #***********************************************************************************************
    # Make monitors for the loads.
    #***********************************************************************************************
    load_object_names = [f"Load.{name}" for name in dss.ActiveCircuit.Loads.AllNames]

    load_monitors = monitor_factory.make_monitors(
        object_names=load_object_names, mode=0, terminal=1
    )

    # Add the monitors to the circuit.
    for monitor in load_monitors:
        dss.Text.Command = monitor.dss_command

    #***********************************************************************************************
    # Make synthetic load profiles.
    #***********************************************************************************************
    # Using a random generator allows us to recreate the same pseudorandom load profiles on each
    # run.
    rng = np.random.default_rng(seed=1337)

    profiles = profile_factory.make_default_gaussian_profiles(
        base_profile_names=["Residential", "Commercial_SM", "Commercial_MD"],
        dss=dss,
        object_names=load_object_names,
        rng=rng
    )

    ckt5_profiles_df = profile_factory.make_default_ckt5_profiles(basepath=basepath)

    # Add the profiles and attach them to each load.
    for profile in profiles:
        for dss_command in profile.dss_commands:
            dss.Text.Command = dss_command

    #***********************************************************************************************
    # Save sythentic load profiles.
    #***********************************************************************************************
    print("Saving synthetic load profiles...")

    profile_df = pd.DataFrame(
        data = [profile.values for profile in profiles],
        columns = range(len(profiles[0].values)),
        index = [profile.element_name for profile in profiles]
    ).T

    pyarrow.feather.write_feather(df=profile_df, dest=f"{basepath}/data/load_profile.feather")

    pyarrow.feather.write_feather(
        df=ckt5_profiles_df, dest=f"{basepath}/data/ckt5_default_load_profile.feather"
    )

    # Free the profiles from memory since it can be a lot.
    del profiles
    del profile_df
    del ckt5_profiles_df

    #***********************************************************************************************
    # Run simulation plan.
    #***********************************************************************************************
    print("Running simulation...")

    # 15-Minute samples for 1 year.
    timestep_count = 35040

    dss.Text.Command = f"set mode=yearly number={timestep_count} stepsize=15m"
    dss.Text.Command = "solve"

    #***********************************************************************************************
    # Make the data.
    #***********************************************************************************************
    print("Making load monitor data...")

    load_voltage_df = data_factory.make_load_voltage_data(dss=dss, monitors=load_monitors)

    #***********************************************************************************************
    # Save data.
    #***********************************************************************************************
    print("Saving data...")

    pyarrow.feather.write_feather(df=load_voltage_df, dest=f"{basepath}/data/load_voltage.feather")
    pyarrow.feather.write_feather(df=metadata_df, dest=f"{basepath}/data/metadata.feather")

    print("...Done!")

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    main()
