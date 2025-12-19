Partial Class StructIQe
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TabstructIQe = Me.Factory.CreateRibbonTab
        Me.grp_ProjectManager = Me.Factory.CreateRibbonGroup
        Me.btnNewProject = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.grp_MailManager = Me.Factory.CreateRibbonGroup
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.grp_QualityManager = Me.Factory.CreateRibbonGroup
        Me.btnSubmit_for_QC = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.grp_DrawingManager = Me.Factory.CreateRibbonGroup
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.grp_TimeManager = Me.Factory.CreateRibbonGroup
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.grp_TaskManager = Me.Factory.CreateRibbonGroup
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Group_Users = Me.Factory.CreateRibbonGroup
        Me.ButtonLogin = Me.Factory.CreateRibbonButton
        Me.Button_Switch_Accounts = Me.Factory.CreateRibbonButton
        Me.grp_ProjectGroupSettings = Me.Factory.CreateRibbonGroup
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.grp_General = Me.Factory.CreateRibbonGroup
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.ButtonHelp = Me.Factory.CreateRibbonButton
        Me.ButtonRefresh = Me.Factory.CreateRibbonButton
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.TabstructIQe.SuspendLayout()
        Me.grp_ProjectManager.SuspendLayout()
        Me.grp_MailManager.SuspendLayout()
        Me.grp_QualityManager.SuspendLayout()
        Me.grp_DrawingManager.SuspendLayout()
        Me.grp_TimeManager.SuspendLayout()
        Me.grp_TaskManager.SuspendLayout()
        Me.Group_Users.SuspendLayout()
        Me.grp_ProjectGroupSettings.SuspendLayout()
        Me.grp_General.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabstructIQe
        '
        Me.TabstructIQe.Groups.Add(Me.grp_ProjectManager)
        Me.TabstructIQe.Groups.Add(Me.grp_MailManager)
        Me.TabstructIQe.Groups.Add(Me.grp_QualityManager)
        Me.TabstructIQe.Groups.Add(Me.grp_DrawingManager)
        Me.TabstructIQe.Groups.Add(Me.grp_TimeManager)
        Me.TabstructIQe.Groups.Add(Me.grp_TaskManager)
        Me.TabstructIQe.Groups.Add(Me.Group_Users)
        Me.TabstructIQe.Groups.Add(Me.grp_ProjectGroupSettings)
        Me.TabstructIQe.Groups.Add(Me.grp_General)
        Me.TabstructIQe.Groups.Add(Me.Group1)
        Me.TabstructIQe.KeyTip = "P"
        Me.TabstructIQe.Label = "PROJECT Assist"
        Me.TabstructIQe.Name = "TabstructIQe"
        '
        'grp_ProjectManager
        '
        Me.grp_ProjectManager.Items.Add(Me.btnNewProject)
        Me.grp_ProjectManager.Items.Add(Me.Button11)
        Me.grp_ProjectManager.Label = "Project"
        Me.grp_ProjectManager.Name = "grp_ProjectManager"
        Me.grp_ProjectManager.Visible = False
        '
        'btnNewProject
        '
        Me.btnNewProject.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnNewProject.KeyTip = "QC"
        Me.btnNewProject.Label = "New"
        Me.btnNewProject.Name = "btnNewProject"
        Me.btnNewProject.OfficeImageId = "NewFolder"
        Me.btnNewProject.ScreenTip = "New Project"
        Me.btnNewProject.ShowImage = True
        Me.btnNewProject.SuperTip = "Create a new project as per your company's guidelines."
        '
        'Button11
        '
        Me.Button11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button11.KeyTip = "QC"
        Me.Button11.Label = "Manage"
        Me.Button11.Name = "Button11"
        Me.Button11.OfficeImageId = "ArrangeByAssignment"
        Me.Button11.ScreenTip = "Manage Projects"
        Me.Button11.ShowImage = True
        Me.Button11.SuperTip = "Manage your projects by viewing, editing, or updating project details as per your" &
    " requirements."
        '
        'grp_MailManager
        '
        Me.grp_MailManager.Items.Add(Me.Button5)
        Me.grp_MailManager.Items.Add(Me.Button6)
        Me.grp_MailManager.Label = "Mail"
        Me.grp_MailManager.Name = "grp_MailManager"
        Me.grp_MailManager.Visible = False
        '
        'Button5
        '
        Me.Button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button5.KeyTip = "QC"
        Me.Button5.Label = "File"
        Me.Button5.Name = "Button5"
        Me.Button5.OfficeImageId = "MailMergeStartMailMergeMenu"
        Me.Button5.ScreenTip = "File Mail(s)"
        Me.Button5.ShowImage = True
        '
        'Button6
        '
        Me.Button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button6.KeyTip = "QC"
        Me.Button6.Label = "Retrieve"
        Me.Button6.Name = "Button6"
        Me.Button6.OfficeImageId = "MailMergeWizard"
        Me.Button6.ScreenTip = "Retrieve Mail(s)"
        Me.Button6.ShowImage = True
        '
        'grp_QualityManager
        '
        Me.grp_QualityManager.Items.Add(Me.btnSubmit_for_QC)
        Me.grp_QualityManager.Items.Add(Me.Button2)
        Me.grp_QualityManager.Items.Add(Me.Button17)
        Me.grp_QualityManager.Label = "Quality"
        Me.grp_QualityManager.Name = "grp_QualityManager"
        Me.grp_QualityManager.Visible = False
        '
        'btnSubmit_for_QC
        '
        Me.btnSubmit_for_QC.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnSubmit_for_QC.KeyTip = "QC"
        Me.btnSubmit_for_QC.Label = "Submit"
        Me.btnSubmit_for_QC.Name = "btnSubmit_for_QC"
        Me.btnSubmit_for_QC.OfficeImageId = "ListToolImport"
        Me.btnSubmit_for_QC.ScreenTip = "Submit"
        Me.btnSubmit_for_QC.ShowImage = True
        Me.btnSubmit_for_QC.SuperTip = "Submit Documents, Drawings or any other type of File for Internal Quality Checks." &
    ""
        '
        'Button2
        '
        Me.Button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button2.KeyTip = "QC"
        Me.Button2.Label = "Status"
        Me.Button2.Name = "Button2"
        Me.Button2.OfficeImageId = "ReviewingPane"
        Me.Button2.ScreenTip = "Status"
        Me.Button2.ShowImage = True
        Me.Button2.SuperTip = "Submit QC Report for the Documents, Drawings or any other type of File after Inte" &
    "rnal Quality Checks."
        '
        'Button17
        '
        Me.Button17.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button17.KeyTip = "QC"
        Me.Button17.Label = "Review"
        Me.Button17.Name = "Button17"
        Me.Button17.OfficeImageId = "MailMergeMatchFields"
        Me.Button17.ScreenTip = "Review"
        Me.Button17.ShowImage = True
        Me.Button17.SuperTip = "Review QC Status/Reports for all Documents, Drawings or any other type of File fo" &
    "r a Project."
        '
        'grp_DrawingManager
        '
        Me.grp_DrawingManager.Items.Add(Me.Button18)
        Me.grp_DrawingManager.Label = "Drawings"
        Me.grp_DrawingManager.Name = "grp_DrawingManager"
        Me.grp_DrawingManager.Visible = False
        '
        'Button18
        '
        Me.Button18.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button18.KeyTip = "QC"
        Me.Button18.Label = "Manage"
        Me.Button18.Name = "Button18"
        Me.Button18.OfficeImageId = "AccessFormModalDialog"
        Me.Button18.ScreenTip = "Manage Drawings"
        Me.Button18.ShowImage = True
        '
        'grp_TimeManager
        '
        Me.grp_TimeManager.Items.Add(Me.Button7)
        Me.grp_TimeManager.Items.Add(Me.Button16)
        Me.grp_TimeManager.Items.Add(Me.Button8)
        Me.grp_TimeManager.Label = "Time Sheet"
        Me.grp_TimeManager.Name = "grp_TimeManager"
        Me.grp_TimeManager.Visible = False
        '
        'Button7
        '
        Me.Button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button7.KeyTip = "QC"
        Me.Button7.Label = "Submit"
        Me.Button7.Name = "Button7"
        Me.Button7.OfficeImageId = "TimeInsert"
        Me.Button7.ScreenTip = "Submit"
        Me.Button7.ShowImage = True
        '
        'Button16
        '
        Me.Button16.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button16.KeyTip = "QC"
        Me.Button16.Label = "Approve"
        Me.Button16.Name = "Button16"
        Me.Button16.OfficeImageId = "KeepBackgroundRemoval"
        Me.Button16.ScreenTip = "Approve"
        Me.Button16.ShowImage = True
        '
        'Button8
        '
        Me.Button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button8.KeyTip = "QC"
        Me.Button8.Label = "View"
        Me.Button8.Name = "Button8"
        Me.Button8.OfficeImageId = "GroupHeaderFooterInsert"
        Me.Button8.ScreenTip = "View"
        Me.Button8.ShowImage = True
        '
        'grp_TaskManager
        '
        Me.grp_TaskManager.Items.Add(Me.Button13)
        Me.grp_TaskManager.Items.Add(Me.Button15)
        Me.grp_TaskManager.Label = "Tasks"
        Me.grp_TaskManager.Name = "grp_TaskManager"
        Me.grp_TaskManager.Visible = False
        '
        'Button13
        '
        Me.Button13.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button13.KeyTip = "QC"
        Me.Button13.Label = "Reminders"
        Me.Button13.Name = "Button13"
        Me.Button13.OfficeImageId = "SetAlerts"
        Me.Button13.ScreenTip = "Reminders"
        Me.Button13.ShowImage = True
        '
        'Button15
        '
        Me.Button15.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button15.KeyTip = "QC"
        Me.Button15.Label = "Company Policies"
        Me.Button15.Name = "Button15"
        Me.Button15.OfficeImageId = "AccessRequests"
        Me.Button15.ScreenTip = "Company Policies"
        Me.Button15.ShowImage = True
        '
        'Group_Users
        '
        Me.Group_Users.Items.Add(Me.ButtonLogin)
        Me.Group_Users.Items.Add(Me.Button_Switch_Accounts)
        Me.Group_Users.Label = "User"
        Me.Group_Users.Name = "Group_Users"
        '
        'ButtonLogin
        '
        Me.ButtonLogin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonLogin.KeyTip = "QC"
        Me.ButtonLogin.Label = "LogIn"
        Me.ButtonLogin.Name = "ButtonLogin"
        Me.ButtonLogin.OfficeImageId = "InsertHighPrivilegeBlock"
        Me.ButtonLogin.ScreenTip = "LogIn to your structIQe Account"
        Me.ButtonLogin.ShowImage = True
        '
        'Button_Switch_Accounts
        '
        Me.Button_Switch_Accounts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button_Switch_Accounts.KeyTip = "QC"
        Me.Button_Switch_Accounts.Label = "Switch"
        Me.Button_Switch_Accounts.Name = "Button_Switch_Accounts"
        Me.Button_Switch_Accounts.OfficeImageId = "RelatedTasksLayoutNow"
        Me.Button_Switch_Accounts.ScreenTip = "Switch within your structIQe Accounts"
        Me.Button_Switch_Accounts.ShowImage = True
        '
        'grp_ProjectGroupSettings
        '
        Me.grp_ProjectGroupSettings.Items.Add(Me.Button12)
        Me.grp_ProjectGroupSettings.Items.Add(Me.Button10)
        Me.grp_ProjectGroupSettings.Items.Add(Me.Button9)
        Me.grp_ProjectGroupSettings.Label = "Project Group"
        Me.grp_ProjectGroupSettings.Name = "grp_ProjectGroupSettings"
        Me.grp_ProjectGroupSettings.Visible = False
        '
        'Button12
        '
        Me.Button12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button12.KeyTip = "QC"
        Me.Button12.Label = "Users"
        Me.Button12.Name = "Button12"
        Me.Button12.OfficeImageId = "InviteAttendees"
        Me.Button12.ScreenTip = "Manage Users"
        Me.Button12.ShowImage = True
        Me.Button12.SuperTip = "Manage users and define their roles within your company"
        '
        'Button10
        '
        Me.Button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button10.KeyTip = "QC"
        Me.Button10.Label = "New"
        Me.Button10.Name = "Button10"
        Me.Button10.OfficeImageId = "UpdateFolderList"
        Me.Button10.ScreenTip = "Create New Project Group"
        Me.Button10.ShowImage = True
        Me.Button10.SuperTip = "Create a New overarching Project Group to organize your projects"
        '
        'Button9
        '
        Me.Button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button9.KeyTip = "QC"
        Me.Button9.Label = "Edit"
        Me.Button9.Name = "Button9"
        Me.Button9.OfficeImageId = "AnimationCustomActionVerbDialog"
        Me.Button9.ScreenTip = "Edit Project Group Settings"
        Me.Button9.ShowImage = True
        Me.Button9.SuperTip = "Modify your Project Group settings/preferences"
        '
        'grp_General
        '
        Me.grp_General.Items.Add(Me.Button4)
        Me.grp_General.Items.Add(Me.ButtonHelp)
        Me.grp_General.Items.Add(Me.ButtonRefresh)
        Me.grp_General.Label = "General"
        Me.grp_General.Name = "grp_General"
        '
        'Button4
        '
        Me.Button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button4.KeyTip = "QC"
        Me.Button4.Label = "About"
        Me.Button4.Name = "Button4"
        Me.Button4.OfficeImageId = "Info"
        Me.Button4.ScreenTip = "About"
        Me.Button4.ShowImage = True
        '
        'ButtonHelp
        '
        Me.ButtonHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonHelp.KeyTip = "QC"
        Me.ButtonHelp.Label = "Help"
        Me.ButtonHelp.Name = "ButtonHelp"
        Me.ButtonHelp.OfficeImageId = "Help"
        Me.ButtonHelp.ScreenTip = "How to use structIQe's softwares. Some handy Tutorial!!"
        Me.ButtonHelp.ShowImage = True
        Me.ButtonHelp.SuperTip = "How to use structIQe's softwares. Some handy Tutorial!!"
        '
        'ButtonRefresh
        '
        Me.ButtonRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonRefresh.KeyTip = "QC"
        Me.ButtonRefresh.Label = "Refresh"
        Me.ButtonRefresh.Name = "ButtonRefresh"
        Me.ButtonRefresh.OfficeImageId = "Recurrence"
        Me.ButtonRefresh.ScreenTip = "Refresh"
        Me.ButtonRefresh.ShowImage = True
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Label = "People"
        Me.Group1.Name = "Group1"
        Me.Group1.Visible = False
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.KeyTip = "QC"
        Me.Button1.Label = "Clients"
        Me.Button1.Name = "Button1"
        Me.Button1.OfficeImageId = "InviteAttendees"
        Me.Button1.ScreenTip = "Manage Users"
        Me.Button1.ShowImage = True
        Me.Button1.SuperTip = "Manage users and define their roles within your company"
        '
        'Button3
        '
        Me.Button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button3.KeyTip = "QC"
        Me.Button3.Label = "Users"
        Me.Button3.Name = "Button3"
        Me.Button3.OfficeImageId = "InviteAttendees"
        Me.Button3.ScreenTip = "Manage Users"
        Me.Button3.ShowImage = True
        Me.Button3.SuperTip = "Manage users and define their roles within your company"
        '
        'StructIQe
        '
        Me.Name = "StructIQe"
        Me.RibbonType = "Microsoft.Excel.Workbook, Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read" &
    ", Microsoft.Outlook.Response.Read, Microsoft.Word.Document"
        Me.Tabs.Add(Me.TabstructIQe)
        Me.TabstructIQe.ResumeLayout(False)
        Me.TabstructIQe.PerformLayout()
        Me.grp_ProjectManager.ResumeLayout(False)
        Me.grp_ProjectManager.PerformLayout()
        Me.grp_MailManager.ResumeLayout(False)
        Me.grp_MailManager.PerformLayout()
        Me.grp_QualityManager.ResumeLayout(False)
        Me.grp_QualityManager.PerformLayout()
        Me.grp_DrawingManager.ResumeLayout(False)
        Me.grp_DrawingManager.PerformLayout()
        Me.grp_TimeManager.ResumeLayout(False)
        Me.grp_TimeManager.PerformLayout()
        Me.grp_TaskManager.ResumeLayout(False)
        Me.grp_TaskManager.PerformLayout()
        Me.Group_Users.ResumeLayout(False)
        Me.Group_Users.PerformLayout()
        Me.grp_ProjectGroupSettings.ResumeLayout(False)
        Me.grp_ProjectGroupSettings.PerformLayout()
        Me.grp_General.ResumeLayout(False)
        Me.grp_General.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabstructIQe As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grp_QualityManager As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnSubmit_for_QC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group_Users As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_ProjectManager As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnNewProject As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_MailManager As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_TimeManager As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_General As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grp_ProjectGroupSettings As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_TaskManager As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button_Switch_Accounts As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents grp_DrawingManager As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonHelp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonRefresh As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property structIQe() As structIQe
        Get
            Return Me.GetRibbon(Of structIQe)()
        End Get
    End Property
End Class
