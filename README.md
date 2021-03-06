# **phase-svd-opendss**

This repository contains the source code for generating the dataset used in the [phase-svd](https://github.com/msk-5s/phase-svd) repository. This code will generate a synthetic voltage magnitude dataset using the [Electric Power Research Institute's (EPRI)](https://www.epri.com/) ckt5 test feeder circuit. This dataset is composed of 35040 voltage magnitude measurements from 1379 loads sampled at 15-minute intervals over a year. Each column name in the dataset will contain the name of the load in the ckt5 circuit. Metadata will also be generated for each load containing the following:
- Name of the load in the ckt5 circuit.
- Phase that the load is connected to.
- Name of the ckt5 loadshape used to synthesize the load profile for the load.
- Name of the secondary distribution transformer that the load is connected to.

The generated dataset will be in the [Apache Arrow Feather](https://arrow.apache.org/docs/python/feather.html) format.

## Requirements
    - Windows
    - Python 3.8+ (64-bit)
    - OpenDSS 9.2.0.1+ (64-bit)
    - See requirements.txt file for the required python packages.
    
## Running
Since the ckt5 test feeder circuit `.dss` files aren't included in this repository, they will need to be copied from your OpenDSS installation folder to the `ckt5-src/` folder. The dataset can then be generated by simply running `run.py`.

> NOTE: During the `Running simulation...` step, the OpenDSS progress bar that pops up may stop responding (as of OpenDSS 9.2.0.1). You can safely ignore this (you just won't see how far along the simulation is) as the simulation is still running. Once the simulation is done, you will have to manually close the OpenDSS progress bar.

## Converting to `.csv` (if desired)
Pandas can be used to convert a `.feather` file to a `.csv` file using the below code:
```
import pyarrow.feather

# A pandas dataframe is returned.
# We are assuming that we are in the repository's root directory.
data_df = pyarrow.feather.read_feather("data/load_voltage.feather")

data_df.to_csv("data/load_voltage.csv")
```
