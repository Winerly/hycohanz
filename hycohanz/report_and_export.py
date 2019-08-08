# -*- coding: utf-8 -*-
#Author: Winerly
'''
Functions in this module correspond to Result>Create Modal Solution Data Report and
Export data have been reported in HFSS.

'''

from __future__ import division, print_function, unicode_literals, absolute_import
import numpy.core.defchararray as npchar

def create_report(oDesign,
                  paraname_array,
                  paravalue_array,
                  x_component,
                  y_component,
                  solution_name,
                  report_name = 'Plot1',
                  report_type = 'Modal Solution Data',
                  display_type = 'Rectangular Plot',
                  domain_type = 'Sweep'
                  ):
    '''

    :param oDesign: pywin32 COMObject
        The HFSS design to which this function is applied.
    :param paraname_array: list of str
        All the names of parameters,include all the project variables and "Freq"
    :param paravalue_array: list of list of str,length equals to paraname_array.
        All the values of parameters,you can input ["All"],["10GHz","11GHz"] and ["Nominal"].
    :param x_component: str
        The parameter name in X axis.
    :param y_component: str or list of str
        The parameter name in Y axis.
        Warning: If the parameter is not like "Freq","Phi" etc,then you must input a list,
                 for example,if you want to set Y axis with S11,the y_component will be
                 ["dB(S(FloquetPort1:1,FloquetPort1:1))"]
    :param report_name:str
        The name of report.
    :param report_type:str
        This parameter is a little fuzzy,in HFSS Scripting Guide,it can be only "Modal S Parameters",
        "Terminal S Parameters","Eigenmode Parameters","Fields","Far Fields","Near Fields" and "Emission Test".
        However,in fact,when I use HFSS's record script to file,I get "Modal Solution Data".
        To normally use,I set its default value with "Modal Solution Data".
    :param display_type:str
        The type of display,for example,"Rectangular Plot", "Polar Plot", "Radiation Pattern","Smith Chart",
        "Data Table", "3D Rectangular Plot", or "3D Polar Plot".
        Warning: It valid values is up to report_type.
    :param solution_name:str
        it represents which solution's result need to report
    :param DomainType:str
        "Sweep" or "Time".
    :return:None
    '''

    paraname_array = npchar.add(paraname_array, ':=')
    para_array = range(2*len(paraname_array))
    para_array[::2] = paraname_array
    para_array[1::2] = paravalue_array

    oReportSetup = oDesign.GetModule("ReportSetup")
    return oReportSetup.CreateReport(report_name,
                                     report_type,
                                     display_type,
                                     solution_name,
                                     ['Domain:=', domain_type],
                                     para_array,
                                     ["X Component:=", x_component,
                                      "Y Component:=", y_component],
                                     [])

def export_to_file(oDesign, report_name, file_name):

    '''

    :param oDesign: pywin32 COMObject
        The HFSS design to which this function is applied.
    :param report_name: str
        The name of the report need to export.
    :param file_path: str
        The name of file(include file path).
    :return: None
    '''

    oReportSetup = oDesign.GetModule("ReportSetup")
    return oReportSetup.ExportToFile(report_name, file_name)