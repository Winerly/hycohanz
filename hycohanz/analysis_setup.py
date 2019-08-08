# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described 
in the HFSS Scripting Guide, Section "Analysis Setup Module Script Commands"

At last count there were 2 functions implemented out of 20.
"""
from __future__ import division, print_function, unicode_literals, absolute_import


def insert_frequency_sweep_linear_discrete(oDesign,
                                           setupname,
                                           sweepname,
                                           startvalue,
                                           stopvalue,
                                           stepsize,
                                           IsEnabled=True,
                                           SetupType="LinearStep",
                                           Type="Discrete",
                                           SaveFields=True,
                                           ExtrapToDC=False):
    """
    Insert an HFSS frequency sweep.

    Warning
    -------
    The API interface for this function is very susceptible to change!  It
    currently only works for Discrete sweeps using Linear Steps.  Contributions
    are encouraged.

    Parameters
    ----------
    oAnalysisSetup : pywin32 COMObject
        The HFSS Analysis Setup Module in which to insert the sweep.
    setupname : string
        The name of the setup to add
    sweepname : string
        The desired name of the sweep
    startvalue : float
        Lowest frequency in Hz.
    stopvalue : float
        Highest frequency in Hz.
    stepsize : flot
        The frequency increment in Hz.
    IsEnabled : bool
        Whether the sweep is enabled.
    SetupType : string
        The type of sweep setup to add.  One of "LinearStep", "LinearCount",
        or "SinglePoints".  Currently only "LinearStep" is supported.
    Type : string
        The type of sweep to perform.  One of "Discrete", "Fast", or
        "Interpolating".  Currently only "Discrete" is supported.
    Savefields : bool
        Whether to save the fields.
    ExtrapToDC : bool
        Whether extrapolation to DC is enabled.

    Returns
    -------
    None

    """
    oAnalysisSetup = oDesign.GetModule("AnalysisSetup")
    return oAnalysisSetup.InsertFrequencySweep(setupname,
                                               ["NAME:" + sweepname,
                                                "IsEnabled:=", IsEnabled,
                                                "SetupType:=", SetupType,
                                                "StartValue:=", str(startvalue) + "GHz",
                                                "StopValue:=", str(stopvalue) + "GHz",
                                                "StepSize:=", str(stepsize) + "GHz",
                                                "Type:=", Type,
                                                "SaveFields:=", SaveFields,
                                                "ExtrapToDC:=", ExtrapToDC])


def insert_analysis_setup(oDesign,
                          Frequency,
                          PortsOnly=False,
                          MaxDeltaS=0.02,
                          Name='Setup1',
                          UseMatrixConv=False,
                          MaximumPasses=6,
                          MinimumPasses=1,
                          MinimumConvergedPasses=1,
                          PercentRefinement=30,
                          IsEnabled=True,
                          BasisOrder=1,
                          UseIterativeSolver=False,
                          DoLambdaRefine=True,
                          DoMaterialLambda=True,
                          SetLambdaTarget=False,
                          Target=0.3333,
                          UseMaxTetIncrease=False,
                          PortAccuracy=2,
                          UseABCOnPort=False,
                          SetPortMinMaxTri=False,
                          EnableSolverDomains=False,
                          SaveRadFieldsOnly=False,
                          SaveAnyFields=True,
                          NoAdditionalRefinementOnImport=False):
    """
    Insert an HFSS analysis setup.
    """
    oAnalysisSetup = oDesign.GetModule("AnalysisSetup")
    oAnalysisSetup.InsertSetup("HfssDriven",
                               ["NAME:" + Name,
                                "Frequency:=", str(Frequency) + "GHz",
                                "PortsOnly:=", PortsOnly,
                                "MaxDeltaS:=", MaxDeltaS,
                                "UseMatrixConv:=", UseMatrixConv,
                                "MaximumPasses:=", MaximumPasses,
                                "MinimumPasses:=", MinimumPasses,
                                "MinimumConvergedPasses:=", MinimumConvergedPasses,
                                "PercentRefinement:=", PercentRefinement,
                                "IsEnabled:=", IsEnabled,
                                "BasisOrder:=", BasisOrder,
                                "UseIterativeSolver:=", UseIterativeSolver,
                                "DoLambdaRefine:=", DoLambdaRefine,
                                "DoMaterialLambda:=", DoMaterialLambda,
                                "SetLambdaTarget:=", SetLambdaTarget,
                                "Target:=", Target,
                                "UseMaxTetIncrease:=", UseMaxTetIncrease,
                                "PortAccuracy:=", PortAccuracy,
                                "UseABCOnPort:=", UseABCOnPort,
                                "SetPortMinMaxTri:=", SetPortMinMaxTri,
                                "EnableSolverDomains:=", EnableSolverDomains,
                                "SaveRadFieldsOnly:=", SaveRadFieldsOnly,
                                "SaveAnyFields:=", SaveAnyFields,
                                "NoAdditionalRefinementOnImport:=", NoAdditionalRefinementOnImport])

    return Name


#Author: Winerly

def insert_frequency_sweep_linear_interpolating(oDesign,
                                                setupname,
                                                sweepname,
                                                startvalue,
                                                stopvalue,
                                                stepsize,
                                                IsEnabled=True,
                                                SetupType="LinearStep",
                                                Type="Interpolating",
                                                SaveFields=False,
                                                SaveRadFields=False,
                                                InterpTolerance=0.5,
                                                InterpMaxSolns=250,
                                                InterpMinSolns=0,
                                                InterpMinSubranges=1,
                                                ExtrapToDC=False,
                                                InterpUseS=True,
                                                InterpUsePortImped=False,
                                                InterpUsePropConst=True,
                                                UseDerivativeConvergence=False,
                                                InterpDerivTolerance=0.2,
                                                UseFullBasis=True,
                                                EnforcePassivity=False):
    oAnalysisSetup = oDesign.GetModule("AnalysisSetup")
    return oAnalysisSetup.InsertFrequencySweep(setupname,
                                               ["NAME:" + sweepname,
                                                "IsEnabled:=", IsEnabled,
                                                "SetupType:=", SetupType,
                                                "StartValue:=", str(startvalue) + "GHz",
                                                "StopValue:=", str(stopvalue) + "GHz",
                                                "StepSize:=", str(stepsize) + "GHz",
                                                "Type:=", Type,
                                                "SaveFields:=", SaveFields,
                                                "SaveRadFields:=", SaveRadFields,
                                                "InterpTolerance:=", InterpTolerance,
                                                "InterpMaxSolns:=", InterpMaxSolns,
                                                "InterpMinSolns:=", InterpMinSolns,
                                                "InterpMinSubranges:=", InterpMinSubranges,
                                                "ExtrapToDC:=", ExtrapToDC,
                                                "InterpUseS:=", InterpUseS,
                                                "InterpUsePortImped:=", InterpUsePortImped,
                                                "InterpUsePropConst:=", InterpUsePropConst,
                                                "UseDerivativeConvergence:=", UseDerivativeConvergence,
                                                "InterpDerivTolerance:=", InterpDerivTolerance,
                                                "UseFullBasis:=", UseFullBasis,
                                                "EnforcePassivity:=", EnforcePassivity])

