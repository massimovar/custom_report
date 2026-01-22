#region Using directives
using System;
using System.Linq;
using FTOptix.HMIProject;
using FTOptix.NetLogic;
using FTOptix.UI;
using UAManagedCore;
using FTOptix.OPCUAServer;
using FTOptix.ODBCStore;
using OpcUa = UAManagedCore.OpcUa;
#endregion

public class CreateUserPanelLogic : BaseNetLogic
{
    [ExportMethod]
    public void CreateUser(string username, string password, string locale, out NodeId result)
    {
        result = NodeId.Empty;

        if (string.IsNullOrEmpty(username))
        {
            ShowMessage(1);
            Log.Error("EditUserDetailPanelLogic", "Cannot create user with empty username");
            return;
        }

        result = GenerateUser(username, password, locale);
    }

    private NodeId GenerateUser(string username, string password, string locale)
    {
        var users = GetUsers();
        if (users == null)
        {
            ShowMessage(2);
            Log.Error("EditUserDetailPanelLogic", "Unable to get users");
            return NodeId.Empty;
        }

        foreach (var child in users.Children.OfType<FTOptix.Core.User>())
        {
            if (child.BrowseName.Equals(username, StringComparison.OrdinalIgnoreCase))
            {
                ShowMessage(10);
                Log.Error("EditUserDetailPanelLogic", "Username already exists");
                return NodeId.Empty;
            }
        }

        var user = InformationModel.MakeObject<FTOptix.Core.User>(username);
        users.Add(user);

        //Apply LocaleId
        if (!string.IsNullOrEmpty(locale))
            user.LocaleId = locale;

        //Apply groups
        ApplyGroups(user);

        //Apply roles
        ApplyRoles(user);

        //Apply password
        var result = Session.ChangePassword(username, password, string.Empty);

        switch (result.ResultCode)
        {
            case FTOptix.Core.ChangePasswordResultCode.Success:
                break;
            case FTOptix.Core.ChangePasswordResultCode.WrongOldPassword:
                //Not applicable
                break;
            case FTOptix.Core.ChangePasswordResultCode.PasswordAlreadyUsed:
                //Not applicable
                break;
            case FTOptix.Core.ChangePasswordResultCode.PasswordChangedTooRecently:
                //Not applicable
                break;
            case FTOptix.Core.ChangePasswordResultCode.PasswordTooShort:
                ShowMessage(6);
                users.Remove(user);
                return NodeId.Empty;
            case FTOptix.Core.ChangePasswordResultCode.UserNotFound:
                //Not applicable
                break;
            case FTOptix.Core.ChangePasswordResultCode.UnsupportedOperation:
                ShowMessage(8);
                users.Remove(user);
                return NodeId.Empty;

        }

        return user.NodeId;
    }

    private void ApplyGroups(FTOptix.Core.User user)
    {
        var groupsPanel = Owner.Get("HorizontalLayout/GroupsAndRoles/GroupsPanel");
        var editable = groupsPanel.GetVariable("Editable");
        var groups = groupsPanel.GetAlias("Groups");
        var panel = groupsPanel.Get("ScrollView/Container");

        UpdateReferences(FTOptix.Core.ReferenceTypes.HasGroup, user, panel, groups, editable);
    }

    private void ApplyRoles(FTOptix.Core.User user)
    {
        var groupsPanel = Owner.Get("HorizontalLayout/GroupsAndRoles/RolesPanel");
        var editable = groupsPanel.GetVariable("Editable");
        var roles = groupsPanel.GetAlias("Roles");
        var panel = groupsPanel.Get("ScrollView/Container");

        UpdateReferences(FTOptix.Core.ReferenceTypes.HasRole, user, panel, roles, editable);
    }

    private static void UpdateReferences(NodeId referenceType, IUANode user, IUANode panel, IUANode rolesOrGroupsAlias, IUAVariable editable)
    {
        if (!editable.Value)
        {
            Log.Debug("EditUserDetailPanelLogic", "User cannot be edited");
            return;
        }

        if (user == null || rolesOrGroupsAlias == null || panel == null)
        {
            Log.Error("EditUserDetailPanelLogic", "User, roles or groups alias or panel not found");
            return;
        }

        var userNode = InformationModel.Get(user.NodeId);
        if (userNode == null)
        {
            Log.Error("EditUserDetailPanelLogic", "User node not found");
            return;
        }

        var referenceCheckBoxes = panel.Refs.GetObjects(OpcUa.ReferenceTypes.HasOrderedComponent, false);

        foreach (var referenceCheckBoxNode in referenceCheckBoxes)
        {
            var role = rolesOrGroupsAlias.Get(referenceCheckBoxNode.BrowseName);
            if (role == null)
            {
                Log.Error("EditUserDetailPanelLogic", "Role or group not found");
                return;
            }

            bool userHasReference = UserHasReference(referenceType, user, role.NodeId);

            if (referenceCheckBoxNode.GetVariable("Checked").Value && !userHasReference)
            {
                Log.Debug("EditUserDetailPanelLogic", $"Adding reference {referenceType} to user {user.NodeId} for role {role.NodeId}");
                userNode.Refs.AddReference(referenceType, role);
            }
            else if (!referenceCheckBoxNode.GetVariable("Checked").Value && userHasReference)
            {
                Log.Debug("EditUserDetailPanelLogic", $"Removing reference {referenceType} from user {user.NodeId} for role {role.NodeId}");
                userNode.Refs.RemoveReference(referenceType, role.NodeId, false);
            }
        }
    }

    private static bool UserHasReference(NodeId referenceType, IUANode user, NodeId groupNodeId)
    {
        if (user == null)
            return false;
        var userGroups = user.Refs.GetObjects(referenceType, false);
        foreach (var userGroup in userGroups)
        {
            if (userGroup.NodeId == groupNodeId)
                return true;
        }
        return false;
    }

    private IUANode GetUsers()
    {
        var pathResolverResult = LogicObject.Context.ResolvePath(LogicObject, "{Users}");
        if (pathResolverResult == null)
            return null;
        if (pathResolverResult.ResolvedNode == null)
            return null;

        return pathResolverResult.ResolvedNode;
    }

    private void ShowMessage(int message)
    {
        var errorMessageVariable = LogicObject.GetVariable("ErrorMessage");
        if (errorMessageVariable != null)
            errorMessageVariable.Value = message;

        delayedTask?.Dispose();

        delayedTask = new DelayedTask(DelayedAction, 5000, LogicObject);
        delayedTask.Start();
    }

    private void DelayedAction(DelayedTask task)
    {
        if (task.IsCancellationRequested)
            return;

        var errorMessageVariable = LogicObject.GetVariable("ErrorMessage");
        if (errorMessageVariable != null)
        {
            errorMessageVariable.Value = 0;
        }
        delayedTask?.Dispose();
    }

    private DelayedTask delayedTask;
}
