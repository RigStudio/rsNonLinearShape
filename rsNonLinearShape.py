# rsNonLinearShape
# @author Roberto Rubio
# @date 2013-08-05
# @file rsNonLinearShape.py

import win32com.client
from win32com.client import constants

Application = win32com.client.Dispatch('XSI.Application').Application
XSIFactory = win32com.client.Dispatch('XSI.Factory')

null = None


##
# Load plugin event.
# @param in_reg: register
# @return Boolean
def XSILoadPlugin(in_reg):
    in_reg.Author = "Roberto Rubio"
    in_reg.Name = "rsNonLinearShape"
    in_reg.Email = "info@rigstudio.com"
    in_reg.URL = "www.rigstudio.com"
    in_reg.Help = 'http://rigstudio.com/rsnonlinearshape/'
    in_reg.Major = 1
    in_reg.Minor = 0
    in_reg.RegisterCommand('rrNonLinearShape', 'rrNonLinearShape')
    in_reg.RegisterMenu(constants.siMenuTbAnimateDeformShapeID, "rsNonLinearShape_Menu", False, False)
    in_reg.RegisterProperty("rsNonLinearShape")
    #RegistrationInsertionPoint - do not remove this line
    return True


##
# Load plugin event.
# @param in_reg: register
# @return Boolean
def XSIUnloadPlugin(in_reg):
    strPluginName = in_reg.Name
    Application.LogMessage(str(strPluginName) + str(" has been unloaded."), constants.siVerbose)
    return True


##
# Menu Init event.
# @param in_ctxt: context
# @return Boolean
def rsNonLinearShape_Menu_Init(in_ctxt):
    oMenu = in_ctxt.Source
    oMenu.AddCallbackItem("rsNonLinearShape", "OnrsNonLinearShapeMenuClicked")
    return True


##
# Parameter creation.
# @param in_ctxt: context
# @return Boolean
def OnrsNonLinearShapeMenuClicked(in_ctxt):
    o_root = Application.ActiveProject.ActiveScene.Root
    Application.AddProp("rsNonLinearShape", o_root)
    return True


##
# Command setup.
# @param in_ctxt: context
# @return Boolean
def rrNonLinearShape_Init(in_ctxt):
    oCmd = in_ctxt.Source
    oCmd.Description = ""
    oCmd.ReturnValue = True
    return True


##
# Create property rsNonLinearShape
# @param None
# @return Boolean
def rrNonLinearShape_Execute():
    o_customProperty = __find_custom_property()
    if o_customProperty is not None:
        Application.DeleteObj(o_customProperty)
    o_customProperty = XSIFactory.CreateObject("rsNonLinearShape")
    Application.InspectObj(o_customProperty, '', 'rsNonLinearShape', constants.siLock, False)
    return True


##
# Create GUI parameters.
# @param in_ctxt: context
# @return Boolean
def rsNonLinearShape_Define(in_ctxt):
    oCustomProperty = in_ctxt.Source
    oCustomProperty.AddParameter2("ControlParameter", constants.siString, "", null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siReadOnly)
    oCustomProperty.AddParameter2("Object", constants.siString, "", null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siReadOnly)
    oCustomProperty.AddParameter2("Parameter_Name", constants.siString, "", null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siKeyable)
    oCustomProperty.AddParameter2("Show_List", constants.siBool, False, null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siKeyable)
    oCustomProperty.AddParameter2("SourceShapeList", constants.siString, "", null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siKeyable)
    oCustomProperty.AddParameter2("Show_Target", constants.siBool, False, null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siKeyable)
    oCustomProperty.AddParameter2("TargetShapeList", constants.siString, "[]", null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siKeyable)
    oCustomProperty.AddParameter2("TargetShapeData", constants.siString, "[]", null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siKeyable)
    oCustomProperty.AddParameter2("SourceShapeData", constants.siString, "[]", null, null, null, null, constants.siClassifUnknown, constants.siPersistable + constants.siKeyable)
    return True


##
# Define GUI layout.
# @param in_ctxt: context
# @return Boolean
def rsNonLinearShape_DefineLayout(in_ctxt):
    oLayout = in_ctxt.Source
    oLayout.Clear()
    oLayout.AddGroup("Object")
    oLayout.AddRow()
    oLayout.AddItem("Object")
    oLayout.AddButton("Pick_Object")
    oLayout.EndRow()
    oLayout.EndGroup()
    oLayout.AddGroup("Control Parameter")
    oLayout.AddGroup("Pick")
    oLayout.AddRow()
    oLayout.AddItem("ControlParameter")
    oLayout.AddButton("Pick_Param")
    oLayout.EndRow()
    oLayout.EndGroup()
    oLayout.AddGroup("New")
    oLayout.AddRow()
    oLayout.AddItem("Parameter_Name")
    oLayout.AddButton("Pick_Control")
    oLayout.EndRow()
    oLayout.EndGroup()
    oLayout.EndGroup()
    oLayout.AddGroup("Shape List")
    oLayout.AddRow()
    oLayout.AddGroup("List")
    oLayout.AddItem("Show_List", "Show")
    o_itemSource = oLayout.AddItem("SourceShapeList", "", constants.siControlListBox)
    o_itemSource.SetAttribute(constants.siUINoLabel, True)
    o_itemSource.WidthPercentage = 50
    o_itemSource.SetAttribute(constants.siUICY, 200)
    o_itemSource.SetAttribute(constants.siUIItems, [])
    oLayout.AddRow()
    oLayout.AddButton("Refresh")
    oLayout.AddButton("Add")
    oLayout.EndRow()
    oLayout.EndGroup()
    oLayout.AddGroup("To Control")
    oLayout.AddItem("Show_Target", "Show")
    o_itemTarget = oLayout.AddItem("TargetShapeList", "", constants.siControlListBox)
    o_itemTarget.SetAttribute(constants.siUINoLabel, True)
    o_itemTarget.WidthPercentage = 50
    o_itemTarget.SetAttribute(constants.siUICY, 200)
    o_itemSource.SetAttribute(constants.siUIItems, [])
    oLayout.AddRow()
    oLayout.AddButton("Remove")
    oLayout.AddButton("Up")
    oLayout.AddButton("Down")
    oLayout.EndRow()
    oLayout.EndGroup()
    oLayout.EndRow()
    oLayout.EndGroup()
    oLayout.AddGroup("Execute")
    oLayout.AddRow()
    oLayout.AddButton("Execute")
    oLayout.AddSpacer(1)
    oLayout.AddButton("Close")
    oLayout.EndRow()
    oLayout.EndGroup()
    return True


##
# OnInit event.
# @param None.
# @return Boolean
def rsNonLinearShape_OnInit():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    Application.SetValue(PPG.Show_Target, 0)
    Application.SetValue(PPG.Show_List, 0)
    PPG.PPGLayout.Item('SourceShapeList').UIItems = []
    PPG.PPGLayout.Item('TargetShapeList').UIItems = []
    Application.SetValue(PPG.Object, "")
    PPG.Refresh
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


##
# OnClosed event.
# @param None.
# @return Boolean
def rsNonLinearShape_OnClosed():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    o_customProperty = __find_custom_property()
    o_paramobject = Application.Dictionary.GetObject("%s.Object" % (o_customProperty), False)
    if o_paramobject != None:
        oParam = o_paramobject.Value
        if oParam != "":
            o_Source = Application.Dictionary.GetObject(oParam, False)
            if o_Source != None:
                Application.SetValue("%s.%s.clustershapecombiner.Mute" % (oParam, o_Source.Type), False, "")
                Application.SetValue("%s.%s.clustershapecombiner.ShowResult" % (oParam, o_Source.Type), True, "")
    Application.DeleteObj(PPG.Inspected(0))
    PPG.Close()
    return True
    Application.SetValue("preferences.scripting.cmdlog", True, "")


##
# Pick Object Method.
# @param None.
# @return Boolean.
def rsNonLinearShape_Pick_Object_OnClicked():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    oParam = PPG.Object
    i_button = -1
    o_param = None
    while i_button != 0:
        l_out = Application.PickObject("Select Control Parameter")
        i_button = l_out[0]
        if i_button == 1:
            o_selection = l_out[2]
            l_shapes = rrSearchShapes(o_selection)
            if l_shapes != []:
                i_button = 0
                o_param = o_selection
            else:
                Application.LogMessage("Object without shapes", 4)
    if o_param is not None:
        Application.SetValue(oParam, o_param)
        PPG.PPGLayout.Item('SourceShapeList').UIItems = l_shapes
        Application.SetValue(PPG.TargetShapeData, str([]))
        PPG.PPGLayout.Item('TargetShapeList').UIItems = []
        Application.SetValue(PPG.SourceShapeData, str(l_shapes))
        Application.SetValue(PPG.SourceShapeList, l_shapes[1], "")
    else:
        Application.SetValue(oParam, "")
        PPG.PPGLayout.Item('SourceShapeList').UIItems = []
    PPG.Refresh()
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


##
# Pick control parameter, the  parameter must be of type double.
# @param None.
# @return Boolean.
def rsNonLinearShape_Pick_Param_OnClicked():
    oParam = PPG.ControlParameter
    i_button = -1
    o_param = None
    while i_button != 0:
        l_out = Application.PickObject("Select Control Parameter")
        i_button = l_out[0]
        if i_button == 1:
            o_selection = l_out[2]
            if o_selection is None or 'Parameter' in o_selection.type:
                if o_selection.ValueType != 5:
                    Application.LogMessage("The control parameter must be float", 4)
                else:
                    o_param = o_selection
                    i_button = 0
    if o_param is not None:
        Application.SetValue("preferences.scripting.cmdlog", False, "")
        Application.SetValue(oParam, o_param)
        Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


##
# Pick control and create parameter, the  parameter must be of type double.
# @param None.
# @return Boolean.
def rsNonLinearShape_Pick_Control_OnClicked():
    oParam = PPG.ControlParameter
    s_Param = PPG.Parameter_Name.Value
    if s_Param != "":
        i_button = -1
        while i_button != 0:
            l_out = Application.PickObject("Select Control Parameter")
            i_button = l_out[0]
            if i_button == 1:
                o_sel = l_out[2]
                Application.SetValue("preferences.scripting.cmdlog", False, "")
                o_custom = Application.Dictionary.GetObject("%s.DisplayInfo_Parameters" % o_sel, False)
                if o_custom == None:
                    Application.AddProp("Custom_parameter_list", o_sel, "", "DisplayInfo_Parameters", "")
                    o_custom = Application.Dictionary.GetObject("%s.DisplayInfo_Parameters" % o_sel, False)
                Application.SIAddCustomParameter(o_custom, s_Param, "siDouble", 0, "", "", "", 2053, "", 1, "", "")
                Application.SetValue(oParam, "%s.%s" % (o_custom.FullName, s_Param))
                i_button = 0
                Application.SetValue("preferences.scripting.cmdlog", True, "")
    else:
        Application.LogMessage("Need a Parameter Name", 2)
    return True


##
# Show List Method. Check to show shapes result.
# @param None.
# @return Boolean.
def rsNonLinearShape_Show_List_OnChanged():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    s_SourceShapeList = PPG.SourceShapeData.Value
    l_SourceShapeList = eval(s_SourceShapeList)
    o_ParamShowList = PPG.Show_List.Value
    if o_ParamShowList:
        Application.SetValue(PPG.Show_Target, 0)
    if len(l_SourceShapeList) > 0:
        rrShowShape()
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


##
# Show Target Method. Check to show shapes result.
# @param None.
# @return Boolean.
def rsNonLinearShape_Show_Target_OnChanged():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    s_TargetShapeList = PPG.TargetShapeData.Value
    l_TargetShapeList = eval(s_TargetShapeList)
    o_ParamTargetList = PPG.Show_Target.Value
    if o_ParamTargetList:
        Application.SetValue(PPG.Show_List, 0)
    if len(l_TargetShapeList) > 0:
        rrShowShape()
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


##
# Source Shape List Method. If the parameter is active visualize the selected shape.
# @param None.
# @return Boolean.
def rsNonLinearShape_SourceShapeList_OnChanged():
    o_ParamShowList = PPG.Show_List.Value
    if o_ParamShowList:
        rrShowShape()
    return True


##
# Target Shape List Method. If the parameter is active visualize the selected shape.
# @param None.
# @return Boolean.
def rsNonLinearShape_TargetShapeList_OnChanged():
    o_ParamTargetList = PPG.Show_Target
    if o_ParamTargetList:
        rrShowShape()
    return True


##
# Refresh Shapes in List To Control.
# @param None.
# @return Boolean.
def rsNonLinearShape_Refresh_OnClicked():
    o_paramShowList = PPG.SourceShapeList.Value
    o_targetShowList = PPG.TargetShapeList.Value
    s_object = PPG.Object.Value
    o_object = Application.Dictionary.GetObject(s_object, False)
    if o_object != None:
        l_shapes = rrSearchShapes(o_object)
        s_TargetShapeList = PPG.TargetShapeData.Value
        l_TargetShapeList = eval(s_TargetShapeList)
        if len(l_shapes) > 0:
            PPG.PPGLayout.Item('SourceShapeList').UIItems = l_shapes
            Application.SetValue(PPG.SourceShapeData, str(l_shapes))
            if str(o_paramShowList) in l_shapes:
                Application.SetValue(PPG.SourceShapeList, o_paramShowList, "")
            else:
                Application.SetValue(PPG.SourceShapeList, l_shapes[1], "")
            rsNonLinearShape_SourceShapeList_OnChanged()
            if len(l_TargetShapeList) > 0:
                l_remove = []
                for z in range(len(l_TargetShapeList)):
                    if str(l_TargetShapeList[z]) not in l_shapes:
                        l_remove.append(l_TargetShapeList[z])
                for s_remove in l_remove:
                    l_TargetShapeList.remove(str(s_remove))
                Application.SetValue(PPG.TargetShapeData, str(l_TargetShapeList))
                PPG.PPGLayout.Item('TargetShapeList').UIItems = l_TargetShapeList
                if str(o_targetShowList) in l_TargetShapeList:
                    Application.SetValue(PPG.TargetShapeList, o_targetShowList, "")
                else:
                    Application.SetValue(PPG.TargetShapeList, l_TargetShapeList[1], "")
                rsNonLinearShape_TargetShapeList_OnChanged()
            PPG.Refresh()
        else:
            Application.SetValue(PPG.TargetShapeData, str([]))
            PPG.PPGLayout.Item('TargetShapeList').UIItems = []
            Application.logMessage("Object Without Shapes", 2)
            PPG.Refresh()
    else:
        Application.logMessage("Need an Object", 4)
    return True


##
# Add Shapes in List To Control.
# @param None.
# @return Boolean.
def rsNonLinearShape_Add_OnClicked():
    s_TargetShapeData = PPG.TargetShapeData.Value
    l_listColor = eval(s_TargetShapeData)
    i_listColor = len(l_listColor)
    oParam = PPG.SourceShapeList
    paramVal = oParam.Value
    o_exist = Application.Dictionary.GetObject(paramVal, False)
    if o_exist != None:
        if str(o_exist) not in l_listColor:
            Application.SetValue("preferences.scripting.cmdlog", False, "")
            l_listColor.append(o_exist.Name)
            l_listColor.append(str(o_exist))
            Application.SetValue(PPG.TargetShapeData, str(l_listColor))
            PPG.PPGLayout.Item('TargetShapeList').UIItems = l_listColor
            if i_listColor == 0:
                Application.SetValue(PPG.TargetShapeList, l_listColor[1], "")
            Application.SetValue("preferences.scripting.cmdlog", True, "")
        else:
            Application.LogMessage("The shape is already in the list", 4)
    PPG.Refresh()
    return True


##
# Remove Shapes in List To Control.
# @param None.
# @return Boolean.
def rsNonLinearShape_Remove_OnClicked():
    s_TargetShapeData = PPG.TargetShapeData.Value
    l_listColor = eval(s_TargetShapeData)
    oParam = PPG.TargetShapeList
    paramVal = oParam.Value
    o_exist = Application.Dictionary.GetObject(paramVal, False)
    if o_exist != None:
        Application.SetValue("preferences.scripting.cmdlog", False, "")
        l_listColor.remove(o_exist.Name)
        l_listColor.remove(str(o_exist))
        Application.SetValue(PPG.TargetShapeData, str(l_listColor))
        PPG.PPGLayout.Item('TargetShapeList').UIItems = l_listColor
        Application.SetValue("preferences.scripting.cmdlog", True, "")
    PPG.Refresh()
    return True


##
# Up Shapes in List To Control.
# @param None.
# @return Boolean.
def rsNonLinearShape_Up_OnClicked():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    s_TargetShapeData = PPG.TargetShapeData.Value
    l_listColor = eval(s_TargetShapeData)
    o_Param = PPG.TargetShapeList
    s_paramVal = o_Param.Value
    i_inListObject = l_listColor.index(s_paramVal)
    i_inListName = i_inListObject - 1
    s_paramValName = l_listColor[i_inListName]
    if i_inListName != 0:
        l_listColor.remove(s_paramValName)
        l_listColor.remove(s_paramVal)
        l_listColor.insert(i_inListName - 2, s_paramValName)
        l_listColor.insert(i_inListObject - 2, s_paramVal)
    else:
        Application.LogMessage("Can not climb", 4)
    Application.SetValue(PPG.TargetShapeData, str(l_listColor))
    PPG.PPGLayout.Item('TargetShapeList').UIItems = l_listColor
    PPG.Refresh()
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


##
# Down Shapes in List To Control.
# @param None.
# @return Boolean.
def rsNonLinearShape_Down_OnClicked():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    s_TargetShapeData = PPG.TargetShapeData.Value
    l_listColor = eval(s_TargetShapeData)
    o_Param = PPG.TargetShapeList
    s_paramVal = o_Param.Value
    i_inListObject = l_listColor.index(s_paramVal)
    i_inListName = i_inListObject - 1
    s_paramValName = l_listColor[i_inListName]
    if i_inListObject != len(l_listColor) - 1:
        l_listColor.remove(s_paramValName)
        l_listColor.remove(s_paramVal)
        l_listColor.insert(i_inListName + 2, s_paramValName)
        l_listColor.insert(i_inListObject + 2, s_paramVal)
    else:
        Application.LogMessage("Can not lower", 4)
    Application.SetValue(PPG.TargetShapeData, str(l_listColor))
    PPG.PPGLayout.Item('TargetShapeList').UIItems = l_listColor
    PPG.Refresh()
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


##
# Make ICE Operator.
# @param None.
# @return Boolean.
def rsNonLinearShape_Execute_OnClicked():
    i_flag = nlsCheck()
    if i_flag == 0:
        Application.LogMessage("Check the log", 2)
        return
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    Application.SetValue("preferences.Interaction.autoinspect", False, "")
    oParam = PPG.Object.Value
    o_Source = Application.Dictionary.GetObject(oParam, False)
    Application.SetValue(PPG.Show_Target, 0)
    Application.SetValue(PPG.Show_List, 0)
    Application.SetValue("%s.%s.clustershapecombiner.Mute" % (oParam, o_Source.Type), True, "")
    Application.SetValue("%s.%s.clustershapecombiner.ShowResult" % (oParam, o_Source.Type), True, "")
    s_param = PPG.ControlParameter.Value
    s_TargetShapeData = PPG.TargetShapeData.Value
    l_listColor = eval(s_TargetShapeData)
    s_precision = "%.3f"
    l_toCluster = []
    for x in range(1, len(l_listColor), 2):
        o_exist = Application.Dictionary.GetObject(l_listColor[x], False)
        if o_exist != None:
            l_elem = list(o_exist.Elements.Array)
            for i_elem in range(len(l_elem[0])):
                l_tmp = []
                [l_tmp.append(abs(l_elem[i_tmp][i_elem])) for i_tmp in range(0, 3)]
                l_tmp = s_precision % sum(l_tmp)
                if l_tmp != s_precision % 0:
                    if i_elem not in l_toCluster:
                        l_toCluster.append(i_elem)
    oCluster = Application.CreateCluster("%s.pnt%s" % (oParam, l_toCluster))
    s_name = "ccmp_rsNonLinearShape_001_MMM"
    i_order = 1
    o_object = Application.Dictionary.GetObject("%s.%s.cls.%s" % (oParam, o_Source.Type, s_name), False)
    while o_object != None:
        i_order = i_order + 1
        s_order = "%03d" % i_order
        s_name = "ccmp_rsNonLinearShape_%s_MMM" % (s_order)
        o_object = Application.Dictionary.GetObject("%s.%s.cls.%s" % (oParam, o_Source.Type, s_name), False)
    Application.SetValue("%s.Name" % (oCluster), s_name, "")
    o_iceTree = Application.Dictionary.GetObject("%s.IceNonLinearShape" % (o_Source), False)
    if o_iceTree == None:
        o_iceTree = Application.ApplyOp("ICETree", o_Source, "siNode", "", "", 0)
        Application.SetValue("%s.Name" % (o_iceTree), "IceNonLinearShape", "")
        o_iceTree = o_iceTree(0)
    l_imputPorts = o_iceTree.InputPorts
    i_imputFlag = 0
    for o_imput in l_imputPorts:
        if not o_imput.IsConnected:
            i_imputFlag = 1
            o_imputPort = o_imput
        else:
            o_connectPort = o_imput
    if not i_imputFlag:
        o_imputPort = Application.AddPortToICENode(o_connectPort, "siNodePortDataInsertionLocationAfter")
    o_GetPointPosition = Application.AddICENode("$XSI_DSPRESETS\\ICENodes\\GetDataNode.Preset", o_iceTree)
    Application.SetValue("%s.reference" % (o_GetPointPosition), "Self.PointPosition", "")
    o_BuildArray = Application.AddICENode("$XSI_DSPRESETS\\ICENodes\\BuildArrayNode.Preset", o_iceTree)
    Application.ConnectICENodes("%s.value1" % (o_BuildArray), "%s.value" % (o_GetPointPosition))
    o_FitBezierCurve = Application.AddICECompoundNode("Fit Bezier Curve", o_iceTree)
    Application.ConnectICENodes("%s.Fit_Points" % (o_FitBezierCurve), "%s.array" % (o_BuildArray))
    Application.AddExpr("%s.T" % (o_FitBezierCurve), s_param, "")
    o_SetPointPosition = Application.AddICECompoundNode("Set Data", o_iceTree)
    Application.SetValue("%s.Reference" % (o_SetPointPosition), "Self.PointPosition", "")
    Application.ConnectICENodes("%s.Value" % (o_SetPointPosition), "%s.Result" % (o_FitBezierCurve))
    o_GetCluster = Application.AddICENode("$XSI_DSPRESETS\\ICENodes\\GetDataNode.Preset", o_iceTree)
    Application.SetValue("%s.reference" % (o_GetCluster), "%s.IsElement" % (oCluster), "")
    o_IfNode = Application.AddICENode("$XSI_DSPRESETS\\ICENodes\\IfNode.Preset", o_iceTree)
    Application.ConnectICENodes("%s.condition" % (o_IfNode), "%s.value" % (o_GetCluster))
    Application.ConnectICENodes("%s.ifTrue" % (o_IfNode), "%s.Execute" % (o_SetPointPosition))
    Application.ConnectICENodes(o_imputPort, "%s.result" % (o_IfNode))
    i_ValueOrder = 2
    for x in range(1, len(l_listColor), 2):
        o_exist = Application.Dictionary.GetObject(l_listColor[x], False)
        o_GetShapePosition = Application.AddICENode("$XSI_DSPRESETS\\ICENodes\\GetDataNode.Preset", o_iceTree)
        Application.SetValue("%s.reference" % (o_GetShapePosition), "%s.positions" % (o_exist), "")
        o_AddNode = Application.AddICENode("$XSI_DSPRESETS\\ICENodes\\AddNode.Preset", o_iceTree)
        Application.ConnectICENodes("%s.value1" % (o_AddNode), "%s.value" % (o_GetPointPosition))
        Application.ConnectICENodes("%s.value2" % (o_AddNode), "%s.value" % (o_GetShapePosition))
        Application.AddPortToICENode("%s.value%s" % (o_BuildArray, (i_ValueOrder - 1)), "siNodePortDataInsertionLocationAfter")
        Application.ConnectICENodes("%s.value%s" % (o_BuildArray, i_ValueOrder), "%s.result" % (o_AddNode))
        i_ValueOrder = i_ValueOrder + 1
    Application.SetValue("preferences.Interaction.autoinspect", True, "")
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


##
# Close and Delete PPG.
# @param None.
# @return Boolean.
def rsNonLinearShape_Close_OnClicked():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    oParam = PPG.Object.Value
    if oParam != "":
        o_Source = Application.Dictionary.GetObject(oParam, False)
        if o_Source != None:
            Application.SetValue("%s.%s.clustershapecombiner.Mute" % (oParam, o_Source.Type), False, "")
            Application.SetValue("%s.%s.clustershapecombiner.ShowResult" % (oParam, o_Source.Type), True, "")
    Application.DeleteObj(PPG.Inspected(0))
    PPG.Close()
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return True


# *******************************************************************
#                            Functions
# *******************************************************************


##
# Search Shapes in object.
# @param o_obj - Object to check.
# @return l_shapes - Shapes list.
def rrSearchShapes(o_obj):
    l_shapes = []
    try:
        l_clusters = o_obj.ActivePrimitive.Geometry.Clusters
        for o_cluster in l_clusters:
            l_clusShape = o_cluster.LocalProperties.Filter("clskey")
            for o_shape in l_clusShape:
                if "ResultClusterKey" not in str(o_shape):
                    l_shapes.append(o_shape.Name)
                    l_shapes.append(str(o_shape))
    except:
        Application.LogMessage("Need a Geometry", 4)
    return l_shapes


##
# Check Parameters to execute.
# @param None.
# @return i_flag - True if the parameters are Ok.
def nlsCheck():
    i_flag = 1
    o_object = PPG.Object.Value
    o_param = PPG.ControlParameter.Value
    s_TargetShapeData = PPG.TargetShapeData.Value
    l_listShape = eval(s_TargetShapeData)
    if o_object != "":
        o_existObject = Application.Dictionary.GetObject(o_object, False)
        if o_existObject == None:
            i_flag = 0
            Application.LogMessage("Object does not exist", 2)
    else:
        i_flag = 0
        Application.LogMessage("Object does not exist", 2)
    if o_param != "":
        o_existParam = Application.Dictionary.GetObject(o_param, False)
        if o_existParam == None:
            i_flag = 0
            Application.LogMessage("Parameter does not exist", 2)
    else:
        i_flag = 0
        Application.LogMessage("Parameter does not exist", 2)
    if len(l_listShape) != 0:
        for o_shape in l_listShape:
            o_existShape = Application.Dictionary.GetObject(o_shape, False)
            if o_existShape == None:
                i_flag = 0
                Application.LogMessage("Shape %s does not exist" % (o_shape), 2)
    else:
        i_flag = 0
    return i_flag


##
# Find Custom Property Method.
# @return o_customProperty - The existing custom property.
# @return (None) If no rsNonLinearShape property exists.
def __find_custom_property():
    c_customProperties = Application.FindObjects(None, "{76332571-D242-11d0-B69C-00AA003B3EA6}")
    for o_customProperty in c_customProperties:
        if o_customProperty.Type == 'rsNonLinearShape':
            return o_customProperty
    return None


##
# Show Shape Result Method. Display the selected shape.
# @param None.
# @return Boolean.
def rrShowShape():
    Application.SetValue("preferences.scripting.cmdlog", False, "")
    oParam = PPG.Object.Value
    o_Source = Application.Dictionary.GetObject(oParam, False)
    o_ParamShowList = PPG.Show_List.Value
    o_ParamTargetList = PPG.Show_Target.Value
    o_paramSource = PPG.SourceShapeList
    s_paramSource = o_paramSource.Value
    o_paramTarget = PPG.TargetShapeList
    s_paramTarget = o_paramTarget.Value
    s_TargetShapeList = PPG.SourceShapeData.Value
    l_TargetShapeList = eval(s_TargetShapeList)
    if o_ParamShowList == 0 and o_ParamTargetList == 0:
        Application.SetValue("%s.%s.clustershapecombiner.Mute" % (oParam, o_Source.Type), False, "")
        Application.SetValue("%s.%s.clustershapecombiner.ShowResult" % (oParam, o_Source.Type), True, "")
    else:
        Application.SetValue("%s.%s.clustershapecombiner.Mute" % (oParam, o_Source.Type), False, "")
        Application.SetValue("%s.%s.clustershapecombiner.ShowResult" % (oParam, o_Source.Type), False, "")
    if o_ParamShowList == 1:
        if s_paramSource != None:
            i_index = l_TargetShapeList.index(s_paramSource) - (l_TargetShapeList.index(s_paramSource) / 2)
            Application.SetValue("%s.%s.clustershapecombiner.SoloIndex" % (oParam, o_Source.Type), i_index, "")
        else:
            Application.SetValue("%s.%s.clustershapecombiner.SoloIndex" % (oParam, o_Source.Type), 0, "")
    if o_ParamTargetList == 1:
        if s_paramTarget != None:
            i_index = l_TargetShapeList.index(s_paramTarget) - (l_TargetShapeList.index(s_paramTarget) / 2)
            Application.SetValue("%s.%s.clustershapecombiner.SoloIndex" % (oParam, o_Source.Type), i_index, "")
        else:
            Application.SetValue("%s.%s.clustershapecombiner.SoloIndex" % (oParam, o_Source.Type), 0, "")
    Application.SetValue("preferences.scripting.cmdlog", True, "")
    return None
