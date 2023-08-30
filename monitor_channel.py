# SPDX-License-Identifier: BSD-3-Clause

"""
This module contains enumerations for channels of a given monitor mode. The enum value is the
respective channel number within the monitor.

The combinations of channels and their respective indices will change depending on the mode and
number of phases of the monitored object. The mode enums in this modules are specific to the ctk5
circuit.

The channel names and indices can be found via the COM interface:
i_monitor = dss.ActiveCircuit.Monitors
i_monitor.Name = "what ever your monitor name is"
channel_names = i_monitor.Header

The enums in this module correspond to the appropriate index within the list returned by
`i_monitor.Header` for a given phase count and mode.
"""

import enum

#===================================================================================================
#===================================================================================================
class OnePhase:
    """
    This class is a namespace for the modes related to single phase objects.
    """
    #===============================================================================================
    #===============================================================================================
    class Mode0(enum.Enum):
        """
        Standard Mode (+0).
        """
        V1 = 1
        VAngle1 = 2
        V2 = 3
        VAngle2 = 4
        I1 = 5
        IAngle1 = 6
        I2 = 7
        IAngle2 = 8

    #===============================================================================================
    #===============================================================================================
    class Mode1(enum.Enum):
        """
        Power Mode (+1).
        """
        S1_KVA = 1
        Ang1 = 2
        S2_KVA = 3
        Ang2 = 4

#===================================================================================================
#===================================================================================================
class ThreePhase:
    """
    This class is a namespace for the modes related to three phase objects.

    The loads and distribution transformers in the ckt5 circuit are single phase. However, the
    substation transformer has three phases.
    """
    #===============================================================================================
    #===============================================================================================
    class Mode0(enum.Enum):
        """
        Standard Mode.
        """
        V1 = 1
        VAngle1 = 2
        V2 = 3
        VAngle2 = 4
        V3 = 5
        VAngle3 = 6
        V4 = 7
        VAngle4 = 8
        I1 = 9
        IAngle1 = 10
        I2 = 11
        IAngle2 = 12
        I3 = 13
        IAngle3 = 14
        I4 = 15
        IAngle4 = 16

    #===============================================================================================
    #===============================================================================================
    class Mode1(enum.Enum):
        """
        Power Mode (+1).
        """
        S1_KVA = 1
        Ang1 = 2
        S2_KVA = 3
        Ang2 = 4
        S3_KVA = 5
        Ang3 = 6
        S4_KVA = 7
        Ang4 = 8
