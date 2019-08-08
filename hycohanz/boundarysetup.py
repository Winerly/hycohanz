# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described 
in the HFSS Scripting Guide, Section "Boundary and Excitation Module Script 
Commands".

At last count there were 4 functions implemented out of 20.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

from hycohanz.design import get_module
from hycohanz.modeler3d import get_face_by_position

def assign_perfect_e(oDesign, boundaryname, facelist, InfGroundPlane=False):
    """
    Create a perfect E boundary.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    boundaryname : str
        The name to give this boundary in the Boundaries tree.
    facelist : list of ints
        The faces to assign to this boundary condition.
    
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    if isinstance(facelist[0],int):
        oBoundarySetupModule.AssignPerfectE(["Name:" + boundaryname, "Faces:=", facelist, "InfGroundPlane:=", InfGroundPlane])
    else:
        oBoundarySetupModule.AssignPerfectE(["Name:" + boundaryname, "Objects:=", facelist, "InfGroundPlane:=", InfGroundPlane])

def assign_radiation(oDesign, 
                     faceidlist, 
                     IsIncidentField=False, 
                     IsEnforcedField=False, 
                     IsFssReference=False, 
                     IsForPML=False,
                     UseAdaptiveIE=False,
                     IncludeInPostproc=True,
                     Name='Rad1'):
    """
    Assign a radiation boundary on the given faces.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list of ints
        The faces to assign to this boundary condition.
    IsIncidentField : bool
        If True, same as checking the "Incident Field" radio button in the 
        Radiation Boundary setup dialog.  Mutually-exclusive with 
        IsEnforcedField
    IsEnforcedField : bool
        If True, same as checking the "Enforced Field" radio button in the 
        Radiation Boundary setup dialog.  Mutually-exclusive with 
        IsIncidentField
    IsFssReference : bool
        If IsEnforcedField is False, is equivalent to checking the 
        "Reference for FSS" check box i the Radiation Boundary setup dialog.
    IsForPML : bool
        Not explored at this time.  Likely use case is when defining a 
        radiation boundary in conjuction with PMLs where the boundary lies on 
        the surface between the PML and the PML base object.
    UseAdaptiveIE : bool
        Not explored at this time.  It is likely that setting this to True is 
        equivalent to selecting the "Model exterior as HFSS-IE domain" check 
        box in the Radiation Boundary setup dialog.
    IncludeInPostproc : bool
        Not explored at this time.  Likely use case is to remove certain 
        boundaries from consideration during certain postprocessing 
        operations, such as when computing the radiation pattern.
    Name : str
        The name to assign the boundary.
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    arg = ["NAME:{0}".format(Name), 
           "Faces:=", faceidlist, 
           "IsIncidentField:=", IsIncidentField, 
           "IsEnforcedField:=", IsEnforcedField, 
           "IsFssReference:=", IsFssReference, 
           "IsForPML:=", IsForPML, 
           "UseAdaptiveIE:=", UseAdaptiveIE, 
           "IncludeInPostproc:=", IncludeInPostproc]
    
    oBoundarySetupModule.AssignRadiation(arg)

def assign_perfect_h(oDesign, boundaryname, facelist):
    """
    Create a perfect H boundary.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    boundaryname : str
        The name to give this boundary in the Boundaries tree.
    facelist : list of ints
        The faces to assign to this boundary condition.
    
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    oBoundarySetupModule.AssignPerfectH(["Name:" + boundaryname, "Faces:=", facelist])

def assign_waveport_multimode(oDesign, 
                              portname, 
                              faceidlist, 
                              Nmodes=1,
                              RenormalizeAllTerminals=True,
                              UseLineAlignment=False,
                              DoDeembed=False,
                              ShowReporterFilter=False,
                              ReporterFilter=[True],
                              UseAnalyticAlignment=False):
    """
    Assign a waveport excitation using multiple modes.
    
    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    portname : str
        Name of the port to create.
    faceidlist : list
        List of face id integers.
    Nmodes : int
        Number of modes with which to excite the port.
        
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    modesarray = ["NAME:Modes"]
    for n in range(0, Nmodes):
        modesarray.append(["NAME:Mode" + str(n + 1),
                           "ModeNum:=", n + 1,
                           "UseIntLine:=", False])

    waveportarray = ["NAME:" + portname, 
                     "Faces:=", faceidlist, 
                     "NumModes:=", Nmodes, 
                     "RenormalizeAllTerminals:=", RenormalizeAllTerminals, 
                     "UseLineAlignment:=", UseLineAlignment, 
                     "DoDeembed:=", DoDeembed, 
                     modesarray, 
                     "ShowReporterFilter:=", ShowReporterFilter, 
                     "ReporterFilter:=", ReporterFilter, 
                     "UseAnalyticAlignment:=", UseAnalyticAlignment]

    oBoundarySetupModule.AssignWavePort(waveportarray)


#Author: Winerly

def assign_master(oDesign,
                  BoundName,
                  FacesIdList,
                  Origin,
                  UPos,
                  ReverseV=False):
    """
        Create a master boundary.

        Parameters
        ----------
        oDesign : pywin32 COMObject
            The HFSS design to which this function is applied.
        BoundName : str
            The name to give this boundary in the Boundaries tree.
        FacesIdList : list of ints
            List of face id integers.
        Origin: list of str,length=3
            The start point of the first coordinate system vector
        UPos: list of str,length=3
            The end point of the first coordinate system vector
        ReverseV: bool
            If False,the second vector is obtained by rotating
            the first vector around the Origin anticlockwise,
            If True,the same thing is done but clockwise.


        Returns
        -------
        None
        """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")

    CoordSysArray = ["NAME:CoordSysVector",
                     "Origin:=", Origin,
                     "UPos:=", UPos]
    MasterArray = ["NAME:" + BoundName,
                   CoordSysArray,
                   "ReverseV:=", ReverseV,
                   "Faces:=", FacesIdList]

    oBoundarySetupModule.AssignMaster(MasterArray)

def assign_slave(oDesign,
                 MasterName,
                 BoundName,
                 FacesIdList,
                 Origin,
                 UPos,
                 Phi = "0deg",
                 Theta = "0deg",
                 UseScanAngles = True,
                 Phase = "0deg",
                 ReverseV=True):
    """
        Create a slave boundary.

        Parameters
        ----------
        oDesign : pywin32 COMObject
            The HFSS design to which this function is applied.
        MasterName: str
            The name of corresponding master boundary.
        BoundName : str
            The name to give this boundary in the Boundaries tree.
        FacesIdList : list of ints
            List of face id integers.
        Origin: list of str,length=3
            The start point of the first coordinate system vector
        UPos: list of str,length=3
            The end point of the first coordinate system vector
        UseScanAngles: bool
            If True, then Phi and Theta should be specified.
            If False, then Phase should be specified.
        Phi,Theta and Phase : str
            The parameters of phase delay of slave boundary.
        ReverseV: bool
            If False,the second vector is obtained by rotating
            the first vector around the Origin anticlockwise,
            If True,the same thing is done but clockwise.


        Returns
        -------
        None
        """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")

    CoordSysArray = ["NAME:CoordSysVector",
                     "Origin:=", Origin,
                     "UPos:=", UPos]
    SlaveArray = ["NAME:" + BoundName,
                  "Faces:=", FacesIdList,
                  CoordSysArray,
                  "ReverseV:=", ReverseV,
                  "Master:=", MasterName,
                  "UseScanAngles:=", UseScanAngles,
                  "Phi:=", Phi,
                  "Theta:=", Theta]

    oBoundarySetupModule.AssignSlave(SlaveArray)

def assign_floquet(oDesign,
                   BoundName,
                   FacesIdList,
                   LatticeAVector_start,
                   LatticeAVector_end,
                   LatticeBVector_start,
                   LatticeBVector_end,
                   NumModes = 2,
                   RenormalizeAllTerminals = True,
                   DoDeembed = False,
                   UseIntLine_1 = False,
                   UseIntLine_2 = False,
                   ShowReporterFilter = False,
                   UseScanAngles = True,
                   Phi = "0deg",
                   Theta = "0deg",

                   IndexM_TE = 0,
                   IndexN_TE = 0,
                   KC2_TE = 0,
                   PropagationState_TE = "Propagating",
                   Attenuation_TE = 0,
                   PolarizationState_TE = "TE",
                   AffectsRefinement_TE = True,

                   IndexM_TM = 0,
                   IndexN_TM = 0,
                   KC2_TM = 0,
                   PropagationState_TM = "Propagating",
                   Attenuation_TM = 0,
                   PolarizationState_TM = "TM",
                   AffectsRefinement_TM = True
                   ):
    """
            Create a Floquet port.

            Parameters
            ----------
            oDesign : pywin32 COMObject
                The HFSS design to which this function is applied.
            BoundName : str
                The name to give this boundary in the Boundaries tree.
            FacesIdList : list of ints
                List of face id integers.
            LatticeAVector_start: list of str,length=3
                The start point of the first coordinate system vector
            LatticeAVector_end: list of str,length=3
                The end point of the first coordinate system vector
            LatticeBVector_start: list of str,length=3
                The start point of the second coordinate system vector
            LatticeBVector_end: list of str,length=3
                The end point of the second coordinate system vector
            Others:
                Default parameters.If you want to learn more about these
                parameters,you can refer to HFSS Scripting Guide.


            Returns
            -------
            None
            """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")

    ModesArray_1 = ["NAME:Mode1",
                   "ModeNum:=", 1,
                   "UseIntLine:=", UseIntLine_1]
    ModesArray_2 = ["NAME:Mode2",
                   "ModeNum:=", 2,
                   "UseIntLine:=", UseIntLine_2]
    ModesArray = ["NAME:Modes",
                 ModesArray_1,
                 ModesArray_2,]

    LatticeArray_A = ["NAME:LatticeAVector",
                      "Start:=", LatticeAVector_start,
                      "End:=", LatticeAVector_end]
    LatticeArray_B = ["NAME:LatticeBVector",
                      "Start:=", LatticeBVector_start,
                      "End:=", LatticeBVector_end]

    ModeList_1 = ["NAME:Mode",
                  "ModeNumber:=", 1,
                  "IndexM:=", IndexM_TE,
                  "IndexN:=", IndexN_TE,
                  "KC2:=", KC2_TE,
                  "PropagationState:=", PropagationState_TE,
                  "Attenuation:=", Attenuation_TE,
                  "PolarizationState:=", PolarizationState_TE,
                  "AffectsRefinement:=", AffectsRefinement_TE]
    ModeList_2 = ["NAME:Mode",
                  "ModeNumber:=", 2,
                  "IndexM:=", IndexM_TM,
                  "IndexN:=", IndexN_TM,
                  "KC2:=", KC2_TM,
                  "PropagationState:=", PropagationState_TM,
                  "Attenuation:=", Attenuation_TM,
                  "PolarizationState:=", PolarizationState_TM,
                  "AffectsRefinement:=", AffectsRefinement_TM]
    ModeList = ["NAME:ModesList",
                ModeList_1,
                ModeList_2]

    FloquetPortArray = ["NAME:" + BoundName,
                        "Faces:=", FacesIdList,
                        "NumModes:=", NumModes,
                        "RenormalizeAllTerminals:=", RenormalizeAllTerminals,
                        "DoDeembed:=", DoDeembed,
                        ModesArray,
                        "ShowReporterFilter:=", ShowReporterFilter,
                        "ReporterFilter:=", [False, False],
                        "UseScanAngles:=", UseScanAngles,
                        "Phi:=", Phi, "Theta:=", Theta,
                        LatticeArray_A, LatticeArray_B,
                        ModeList]

    oBoundarySetupModule.AssignFloquetPort(FloquetPortArray)

def assign_lumpedRLC(oDesign,
                     BoundName,
                     Resistance,
                     FacesIdList,
                     StartPoint,
                     EndPoint,
                     UseResist = True,
                     UseInduct = False,
                     UseCap = False,
                     Inductance = 0,
                     Capacitance = 0
                     ):
    """
    :param oDesign: pywin32 COMObject
                The HFSS design to which this function is applied.
    :param BoundName: str
                The name to give this boundary in the Boundaries tree.
    :param Resistance: int
                The resistance value.
    :param FacesIdList: list of ints
                List of face id integers.
    :param StartPoint: list of strs(len = 3)
                The coordinates of current line's start point.
    :param EndPoint: list of strs(len = 3)
                The coordinates of current line's end point.
    :param UseResist: bool
                Resistance is valid or not.
    :param UseInduct: bool
                Inductance is valid or not.
    :param UseCap:bool
                Capacitance is valid or not.
    :param Inductance:
                The inductance value.
    :param Capacitance:
                The capacitance value.
    :return: None
    """

    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")

    if UseResist:
        if isinstance(Resistance, int):
            res = ["UseResist:=", True,
                   "Resistance:=", str(Resistance) + "ohm",]
        else:
            res = ["UseResist:=", True,
                   "Resistance:=", Resistance]
    else:
        res = ["UseResist:=", False]
    if UseInduct:
        ind = ["UseInduct:=", True,
               "Inductance:=", str(Inductance) + "nH"]
    else:
        ind = ["UseInduct:=", False]
    if UseCap:
        cap = ["UseCap:=", True,
               "Capacitance:=", str(Capacitance) + "pF"]
    else:
        cap = ["UseCap:=", False]
    RLC = res + ind + cap

    CurrentLineArray = ["NAME:CurrentLine",
                        "Start:=", StartPoint,
                        "End:=", EndPoint]

    LumpedRLCArray = ["NAME:" + BoundName,
                      "Objects:=", FacesIdList,
                      CurrentLineArray] + RLC

    oBoundarySetupModule.AssignLumpedRLC(LumpedRLCArray)

def assign_box_master_and_slave(oDesign,
                                oEditor,
                                box):

    p = box.get_face_point('front')
    e = box.get_face_edge('front')
    master_id_1 = get_face_by_position(oEditor, box.name, p[0], p[1], p[2])
    assign_master(oDesign, "Master1", [master_id_1],
                       e[0], e[1])

    p = box.get_face_point('rear')
    e = box.get_face_edge('rear')
    slave_id_1 = get_face_by_position(oEditor, box.name, p[0], p[1], p[2])
    assign_slave(oDesign, "Master1", "Slave1", [slave_id_1],
                      e[0], e[1])

    p = box.get_face_point('right')
    e = box.get_face_edge('right')
    master_id_2 = get_face_by_position(oEditor, box.name, p[0], p[1], p[2])
    assign_master(oDesign, "Master2", [master_id_2],
                       e[0], e[1])

    p = box.get_face_point('left')
    e = box.get_face_edge('left')
    slave_id_2 = get_face_by_position(oEditor, box.name, p[0], p[1], p[2])
    assign_slave(oDesign, "Master2", "Slave2", [slave_id_2],
                      e[0], e[1])

def assign_box_floquet(oDesign,
                       oEditor,
                       box):
    p = box.get_face_point('up')
    e = box.get_face_edge('up')
    floquet_id = get_face_by_position(oEditor, box.name, p[0], p[1], p[2])
    assign_floquet(oDesign, "Floquet1", [floquet_id],
                        e[0], e[1], e[2], e[3])
    p = box.get_face_point('down')
    e = box.get_face_edge('down')
    floquet_id = get_face_by_position(oEditor, box.name, p[0], p[1], p[2])
    assign_floquet(oDesign, "Floquet2", [floquet_id],
                        e[0], e[1], e[2], e[3])