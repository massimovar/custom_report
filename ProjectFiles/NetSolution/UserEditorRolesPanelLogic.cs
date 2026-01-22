#region Using directives
using FTOptix.HMIProject;
using FTOptix.NetLogic;
using FTOptix.UI;
using UAManagedCore;
using FTOptix.OPCUAServer;
using FTOptix.ODBCStore;
using OpcUa = UAManagedCore.OpcUa;
#endregion

public class UserEditorRolesPanelLogic : BaseNetLogic
{
    public override void Start()
    {
        userVariable = Owner.GetVariable("User");
        editable = Owner.GetVariable("Editable");

        userVariable.VariableChange += UserVariable_VariableChange;
        editable.VariableChange += Editable_VariableChange;

        UpdateRolesAndUser();

        BuildUIRoles();
        if (editable.Value)
            SetCheckedValues();
    }

    private void Editable_VariableChange(object sender, VariableChangeEventArgs e)
    {
        UpdateRolesAndUser();
        BuildUIRoles();

        if (e.NewValue)
            SetCheckedValues();
    }

    private void UserVariable_VariableChange(object sender, VariableChangeEventArgs e)
    {
        UpdateRolesAndUser();
        if (editable.Value)
            SetCheckedValues();
        else
            BuildUIRoles();
    }

    private void UpdateRolesAndUser()
    {
        if (userVariable.Value.Value != null)
            user = InformationModel.Get(userVariable.Value);

        roles = LogicObject.GetAlias("Roles");
    }

    private void BuildUIRoles()
    {
        if (roles == null)
            return;

        panel?.Delete();

        panel = InformationModel.MakeObject<ColumnLayout>("Container");
        panel.HorizontalAlignment = HorizontalAlignment.Stretch;
        panel.VerticalAlignment = VerticalAlignment.Stretch;
        panel.TopMargin = 0;

        foreach (var role in roles.Children)
        {
            Panel roleUiObject = null;

            if (editable.Value)
            {
                roleUiObject = InformationModel.MakeObject<Panel>(role.BrowseName, custom_report.ObjectTypes.GroupCheckbox);
            }
            else if (UserHasRole(role.NodeId))
            {
                roleUiObject = InformationModel.MakeObject<Panel>(role.BrowseName, custom_report.ObjectTypes.GroupLabel);
            }

            if (roleUiObject == null)
                continue;

            roleUiObject.HorizontalAlignment = HorizontalAlignment.Stretch;
            roleUiObject.VerticalAlignment = VerticalAlignment.Top;
            roleUiObject.TopMargin = 0;

            roleUiObject.GetVariable("Group").Value = role.NodeId;

            panel.Add(roleUiObject);
            panel.Height += roleUiObject.Height;
        }

        var scrollView = Owner.Get("ScrollView");
        scrollView?.Add(panel);
    }

    private void SetCheckedValues()
    {
        if (roles == null)
            return;

        if (panel == null)
            return;

        var roleCheckBoxes = panel.Refs.GetObjects(OpcUa.ReferenceTypes.HasOrderedComponent, false);

        foreach (var roleCheckBoxNode in roleCheckBoxes)
        {
            var role = roles.Get(roleCheckBoxNode.BrowseName);
            roleCheckBoxNode.GetVariable("Checked").Value = UserHasRole(role.NodeId);
        }
    }

    private bool UserHasRole(NodeId roleNodeId)
    {
        if (user == null)
            return false;
        var userRoles = user.Refs.GetObjects(FTOptix.Core.ReferenceTypes.HasRole, false);
        foreach (var userRole in userRoles)
        {
            if (userRole.NodeId == roleNodeId)
                return true;
        }
        return false;
    }

    private IUAVariable userVariable;
    private IUAVariable editable;

    private IUANode roles;
    private IUANode user;
    private ColumnLayout panel;
}
