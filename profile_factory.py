# SPDX-License-Identifier: MIT

"""
The module contains functions for making load profiles.
"""

from typing import Any, List, NamedTuple, Sequence
from nptyping import NDArray

import win32com.client

from rich.progress import track

import numpy as np
import pandas as pd
import scipy.interpolate

#===================================================================================================
#===================================================================================================
class Profile(NamedTuple):
    """
    A tuple for load profile information.
    
    Attributes
    ----------
    dss_commands : list of str
        The OpenDSS commands to create and add the load profile.
    element_name : str
        The name of the load element.
    values : numpy.ndarray of float, (n_timestep,)
        The time series values of the synthetic load profile.
    """
    dss_commands: Sequence[str]
    element_name: str
    values: NDArray[(Any,), float]

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def make_default_ckt5_profiles(basepath: str) -> pd.DataFrame:
    """
    Make the default ckt5 load profiles.

    Parameters
    ----------
    basepath : str
        The basepath for the directory that the main script this is ran from.

    Returns
    -------
    pandas.DataFrame
        The load profiles.
    """
    profiles = {
        "Residential": None,
        "Commercial_SM": None,
        "Commercial_MD": None
    }
    
    with open(f"{basepath}/src/Loadshapes_ckt5.dss", "r", encoding="utf-8") as file:
        for (key, line) in zip(profiles.keys(), file):
            # Get the 'mult' array string.
            array_string = line.split("=")[3]
            array_string = array_string.replace("(", "").replace(")", "")
            
            values = array_string.split(",")
            values = [value.replace(" ", "") for value in values]
            values = np.array([float(value) for value in values])

            profiles[key] = values

    profile_df = pd.DataFrame(data=profiles)

    return profile_df

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def make_default_gaussian_profiles(
    base_profile_names: Sequence[str], dss: win32com.client.CDispatch,
    object_names: Sequence[str], rng: np.random.Generator) -> List[Profile]:
    """
    Make real power load profiles using a base profile and adding gaussian white noise to it.

    Since each element is already assigned a yearly base profile in the ckt5 circuit, they will be
    assigned a synthetic profile built from their base profile.

    Parameters
    ----------
    base_profile_names : list of str
        The name of the base profiles to build the random walk models from.
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.
    object_names : list of str
        The name of the OpenDSS objects to generate a load profile for.
    rng : numpy.random.Generator
        The random generator to use.

    Returns
    -------
    list of profile_factory.Profile
        The new profiles for each object by `base_profile_name`.
    """
    generators = {
        name: _profile_generator(base_profile_name=name, dss=dss, rng=rng)
    for name in base_profile_names}

    i_loads = dss.ActiveCircuit.Loads
    profiles = []

    for object_name in track(object_names, "Making synthetic profiles..."):
        element_name = object_name.split(".")[-1]
        i_loads.Name = element_name
        profile_name = f"{element_name}_profile"

        new_profile = next(generators[i_loads.Yearly])

        # Convert the new profile to a string of an OpenDSS array.
        values = [str(value) for value in new_profile]
        values = "[" + ", ".join(values) + "]"

        # Make an OpenDSS array for the hours.
        hours = [str(0.25 * i) for i in range(len(new_profile))]
        hours = "[" + ", ".join(hours) + "]"

        # NOTE: It's important to specify `npts` before any array properties (`hour`, `mult`, etc).
        # Not doing this will cause OpenDSS 9.2.0.1 to have a memory access violation when running
        # the simulation.
        dss_commands = [
            f"new Loadshape.{profile_name} npts={len(new_profile)} hour={hours} mult={values}",
            f"{object_name}.yearly={profile_name}"
        ]

        profile = Profile(dss_commands=dss_commands, element_name=element_name, values=new_profile)

        profiles.append(profile)

    return profiles

#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
def _profile_generator(
    base_profile_name: str, dss: win32com.client.CDispatch, rng: np.random.Generator)\
    -> NDArray[(Any,), float]:
    """
    Generates a sythetic load profile by adding gaussian white noise to some base profile.

    Parameters
    ----------
    base_profile_name : str
        The name of the base profile to add gaussian white noise to.
    dss : win32com.client.CDispatch
        The OpenDSSEngine object.
    rng : numpy.random.Generator
        The random generator to use.

    Yields
    ------
    numpy.ndarary of float, (n_timestep,)
        A new sythentic load profile generated from a base profile plus some gaussian white noise.
    """
    i_loadshapes = dss.ActiveCircuit.LoadShapes
    i_loadshapes.Name = base_profile_name

    # Use linear interpolation to convert the 1-hour period based load profile to a 15-minute based
    # load profile. This converts the 8760 samples (24 hours * 365 days) to 35037 samples.
    step = 0.25
    temp = np.array(i_loadshapes.Pmult)

    # Interpolation will only give us 35037 samples. The last three values are filled in via linear
    # extrapolation to get to 35040 samples.
    base = scipy.interpolate.interp1d(
        x = np.arange(len(temp)), y = temp, kind="linear", fill_value="extrapolate"
    )(np.arange(start=0, stop=len(temp), step=step))

    while True:
        new_profile = base + rng.normal(loc=0, scale=0.1, size=base.size)

        # Scale the profile to be in the range [0, 1].
        new_profile = np.interp(x=new_profile, xp=(new_profile.min(), new_profile.max()), fp=(0, 1))

        yield new_profile
