# -*- coding: utf-8 -*-
#Author: Winerly
'''
Functions in this module correspond more or less to the functions described
in the HFSS Scripting Guide, Section "Optimetrics Module Script Commands"
'''

from __future__ import division, print_function, unicode_literals, absolute_import

def insert_optimetrics_setup(oDesign,
                             setup_name,
                             Variable,
                             Data,
                             op_name = 'ParametricSetup1',
                             IsEnabled = True,
                             SaveFields = False,
                             CopyMesh = False,
                             OffsetF1 = False,
                             Synchronize = 0):

    '''
    Insert an HFSS optimetrics setup.

    Parameters
    -------
    :param oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    :param setup_name: str
        The name of setup corresponding to that optimetrics setup.
    :param Variable: list of str
        The names of parameters need to sweep.
    :param Data: list of str,its length is the same as Variable.
        The type of sweep and start,end,step of corresponding parameters,
        for example,["LIN 12mm 17mm 2.5mm", "LIN 8mm 12mm 2mm"],
        particularly,it can be more than one type.
        ["2mm LINC 1mm 10mm 15 DEC 1mm 10mm 10 OCT 1mm 10mm 10 ESTP 1mm 10mm 10"]
        A single value is a single value type.
        LINC is "linear count".
        DEC is "decade count".
        OCT is "octave count".
        ESTP is "exponential count".
    :param op_name:
        The name of this optimetrics setup.
    :param IsEnabled: bool
    :param SaveFields: bool
    :param CopyMesh: bool
    :param OffsetF1: bool
    :param Synchronize: int
        Others unimportant parameters.
    :return: None
    -------

    '''

    oOptimetricsSetup = oDesign.GetModule("Optimetrics")

    V = ["NAME:Sweeps"]
    for i in range(len(Variable)):
        V.append(["NAME:SweepDefinition",
                  "Variable:=", r'$'+Variable[i],
                  "Data:=", Data[i],
                  "OffsetF1:=", OffsetF1,
                  "Synchronize:=", Synchronize])

    return oOptimetricsSetup.InsertSetup("OptiParametric",
                                         ["NAME:" + op_name,
                                          "IsEnabled:=", IsEnabled,
                                          ["NAME:ProdOptiSetupData",
                                          "SaveFields:=", SaveFields,
                                          "CopyMesh:=", CopyMesh],
                                          ["NAME:StartingPoint"],
                                          "Sim. Setups:=",
                                          [setup_name],
                                          V,
                                          ["NAME:Sweep Operations"],
                                          ["NAME:Goals"]])
