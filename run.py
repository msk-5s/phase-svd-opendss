# SPDX-License-Identifier: BSD-3-Clause

"""
This script generates a synthetic dataset from EPRI's ckt5 test feeder circuit.
"""

import os
import win32com.client

from rich.progress import track

import numpy as np
import pandas as pd
import pyarrow.feather

import data_factory
import label_factory
import monitor_channel
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
    # Make the labels.
    #***********************************************************************************************
    # NOTE: It's important to make the labels before assigning the synthetic load profiles to the
    # loads in the circuit. We want to have the names of the original load profiles before changing
    # them so that we will know what each type of load is (Residential, Commercial_SM,
    # Commercial_MD).
    print("Making labels...")

    labels_load_df = label_factory.make_load_labels(dss=dss)
    labels_xfmr_df = label_factory.make_xfmr_labels(dss=dss, labels_load_df=labels_load_df)

    #***********************************************************************************************
    # Make monitors for the loads.
    #***********************************************************************************************
    load_object_names = [f"Load.{name}" for name in dss.ActiveCircuit.Loads.AllNames]

    load_monitors = monitor_factory.make_monitors(
        object_names=load_object_names, mode=0, terminal=1
    )

    # Add the monitors to the circuit.
    for monitor in track(load_monitors, "Assigning monitors to loads..."):
        dss.Text.Command = monitor.dss_command

    #***********************************************************************************************
    # Make monitors for the transformers.
    #***********************************************************************************************
    xfmr_sub_name = "MDV_SUB_1".lower()
    xfmr_sub_object_names = [f"Transformer.{xfmr_sub_name}"]

    xfmr_dist_object_names = [
        f"Transformer.{name}"
    for name in dss.ActiveCircuit.Transformers.AllNames if name != xfmr_sub_name]

    xfmr_dist_primary_monitors = monitor_factory.make_monitors(
        object_names=xfmr_dist_object_names, mode=0, terminal=1
    )

    xfmr_dist_secondary_monitors = monitor_factory.make_monitors(
        object_names=xfmr_dist_object_names, mode=0, terminal=2
    )

    xfmr_sub_primary_monitors = monitor_factory.make_monitors(
        object_names=xfmr_sub_object_names, mode=0, terminal=1
    )

    xfmr_sub_secondary_monitors = monitor_factory.make_monitors(
        object_names=xfmr_sub_object_names, mode=0, terminal=2
    )

    # Add the distribution xfmr monitors to the circuit.
    xfmr_dist_monitors = zip(xfmr_dist_primary_monitors, xfmr_dist_secondary_monitors)
    for (primary_monitor, secondary_monitor) in track(
        xfmr_dist_monitors, "Assigning monitors to the distribution xfmrs..."
    ):
        dss.Text.Command = primary_monitor.dss_command
        dss.Text.Command = secondary_monitor.dss_command

    # Add the substation xfmr monitors to the circuit.
    xfmr_sub_monitors = zip(xfmr_sub_primary_monitors, xfmr_sub_secondary_monitors)
    for (primary_monitor, secondary_monitor) in track(
        xfmr_sub_monitors, "Assigning monitors to the substation xfmrs..."
    ):
        dss.Text.Command = primary_monitor.dss_command
        dss.Text.Command = secondary_monitor.dss_command

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
    for profile in track(profiles, "Assigning synthetic profiles to loads..."):
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
    # Make the data for the load and distribution transformers.
    #***********************************************************************************************
    print("Making monitor data...")

    monitors = [load_monitors, xfmr_dist_primary_monitors, xfmr_dist_secondary_monitors]

    key_map = {
        "voltage-magnitude": [
            "load-voltage-magnitude",
            "xfmr-distribution-primary-voltage-magnitude",
            "xfmr-distribution-secondary-voltage-magnitude"
        ],
        "voltage-angle": [
            "load-voltage-angle",
            "xfmr-distribution-primary-voltage-angle",
            "xfmr-distribution-secondary-voltage-angle"
        ],
        "current-magnitude": [
            "load-current-magnitude",
            "xfmr-distribution-primary-current-magnitude",
            "xfmr-distribution-secondary-current-magnitude"
        ],
        "current-angle": [
            "load-current-angle",
            "xfmr-distribution-primary-current-angle",
            "xfmr-distribution-secondary-current-angle"
        ]
    }

    voltage_magnitude_dfs = {
        key: monitor_factory.make_monitor_data(
            channel=monitor_channel.OnePhase.Mode0.V1.value, dss=dss, monitors=monitors
        )
    for (key, monitors) in zip(key_map["voltage-magnitude"], monitors)}

    voltage_angle_dfs = {
        key: monitor_factory.make_monitor_data(
            channel=monitor_channel.OnePhase.Mode0.VAngle1.value, dss=dss, monitors=monitors
        )
    for (key, monitors) in zip(key_map["voltage-angle"], monitors)}

    current_magnitude_dfs = {
        key: monitor_factory.make_monitor_data(
            channel=monitor_channel.OnePhase.Mode0.I1.value, dss=dss, monitors=monitors
        )
    for (key, monitors) in zip(key_map["current-magnitude"], monitors)}

    current_angle_dfs = {
        key: monitor_factory.make_monitor_data(
            channel=monitor_channel.OnePhase.Mode0.IAngle1.value, dss=dss, monitors=monitors
        )
    for (key, monitors) in zip(key_map["current-angle"], monitors)}

    #***********************************************************************************************
    # Make the data for the substation transformer.
    #***********************************************************************************************
    xfmr_channel_map = {
        "voltage-magnitude": [
            monitor_channel.ThreePhase.Mode0.V1.value,
            monitor_channel.ThreePhase.Mode0.V2.value,
            monitor_channel.ThreePhase.Mode0.V3.value
        ],
        "voltage-angle": [
            monitor_channel.ThreePhase.Mode0.VAngle1.value,
            monitor_channel.ThreePhase.Mode0.VAngle2.value,
            monitor_channel.ThreePhase.Mode0.VAngle3.value
        ],
        "current-magnitude": [
            monitor_channel.ThreePhase.Mode0.I1.value,
            monitor_channel.ThreePhase.Mode0.I2.value,
            monitor_channel.ThreePhase.Mode0.I3.value
        ],
        "current-angle": [
            monitor_channel.ThreePhase.Mode0.IAngle1.value,
            monitor_channel.ThreePhase.Mode0.IAngle2.value,
            monitor_channel.ThreePhase.Mode0.IAngle3.value
        ]
    }

    # Make a numpy array for each individual substation channel.
    temp_xfmr_primary_map = {
        key: [
            monitor_factory.make_monitor_data(
                channel=channel, dss=dss, monitors=xfmr_sub_primary_monitors
            ).to_numpy(dtype=float).reshape(-1,)
        for channel in channels]
    for (key, channels) in xfmr_channel_map.items()}

    temp_xfmr_secondary_map = {
        key: [
            monitor_factory.make_monitor_data(
                channel=channel, dss=dss, monitors=xfmr_sub_secondary_monitors
            ).to_numpy(dtype=float).reshape(-1,)
        for channel in channels]
    for (key, channels) in xfmr_channel_map.items()}

    # Column stack the arrays so they form an N x 3 matrix of measurements for each phase.
    xfmr_primary_data = {
        key: np.column_stack(channel_data_list)
    for (key, channel_data_list) in temp_xfmr_primary_map.items()}

    xfmr_secondary_data = {
        key: np.column_stack(channel_data_list)
    for (key, channel_data_list) in temp_xfmr_secondary_map.items()}

    # Make the final dataframes.
    column_names = [f"{xfmr_sub_name};{phase}" for phase in ["a", "b", "c"]]
    xfmr_sub_primary_dfs = {
        key: pd.DataFrame(data=data, columns=column_names)
    for (key, data) in xfmr_primary_data.items()}

    xfmr_sub_secondary_dfs = {
        key: pd.DataFrame(data=data, columns=column_names)
    for (key, data) in xfmr_secondary_data.items()}

    #***********************************************************************************************
    # Save data.
    #***********************************************************************************************
    print("Saving data...")

    data_dfs_list = [
        voltage_magnitude_dfs,
        voltage_angle_dfs,
        current_magnitude_dfs,
        current_angle_dfs
    ]

    _ = [
        [
            pyarrow.feather.write_feather(df=data_df, dest=f"{basepath}/data/{key}.feather")
        for (key, data_df) in data_dfs.items()]
    for data_dfs in data_dfs_list]

    _ = {
        pyarrow.feather.write_feather(
            df=data_df, dest=f"{basepath}/data/xfmr-substation-primary-{key}.feather")
    for (key, data_df) in xfmr_sub_primary_dfs.items()}

    _ = {
        pyarrow.feather.write_feather(
            df=data_df, dest=f"{basepath}/data/xfmr-substation-secondary-{key}.feather")
    for (key, data_df) in xfmr_sub_secondary_dfs.items()}

    pyarrow.feather.write_feather(df=labels_load_df, dest=f"{basepath}/data/load-labels.feather")
    pyarrow.feather.write_feather(df=labels_xfmr_df, dest=f"{basepath}/data/xfmr-labels.feather")

    print("...Done!")

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    main()
