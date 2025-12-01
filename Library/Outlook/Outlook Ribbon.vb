Imports System.Diagnostics
Imports System.Net
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports PROJECT_Assist_Common_Library
Imports structIQe_Common_Library

Public Class StructIQe

    Private DeferTimer As Windows.Forms.Timer
    Private StartupSw As New Stopwatch()

    Private Sub StructIQe_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        TabstructIQe.Label = SettingsHelper.App_Name

        StartupSw.Start()

        Try
            ' 3) Defer all heavy/fragile work (MAPI folder access, disk, network)
            DeferTimer = New Windows.Forms.Timer() With {.Interval = 2000}  ' 2s after Outlook is up
            AddHandler DeferTimer.Tick, AddressOf DeferredInit
            DeferTimer.Start()

        Catch ex As System.Exception
            ' Never let Startup throw
        Finally
            StartupSw.Stop()
        End Try

    End Sub

    Private Sub DeferredInit(sender As Object, e As EventArgs)
        DeferTimer.Stop()
        RemoveHandler DeferTimer.Tick, AddressOf DeferredInit
        DeferTimer.Dispose()
        Dim sw As Stopwatch = Stopwatch.StartNew()

        Try

            Refresh_Ribbon_As_per_permissions()

        Catch ex As System.Exception
        Finally
            sw.Stop()
        End Try
    End Sub

    Async Sub Refresh_Ribbon_As_per_permissions()

        If SupabaseHelper.IsOnline() = False Then
            Hide_all_buttons()
            Exit Sub
        End If

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Dim access = Await SupabaseHelper.TryAutoLoginAsync()

        If access Is Nothing Then
            Hide_all_buttons()
            Dim res = SharedRibbonButtons.User_Login()
            If res Is Nothing Then
                System.Windows.Forms.Cursor.Current = Cursors.Default
                Exit Sub
            ElseIf res = True Then
                access = SharedRibbonButtons.CurrentAccess
                SharedRibbonButtons.Update_app_name()
            Else
                access = Nothing
            End If

        End If


        Reset_ribbon_from_Database(access)

        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

    Sub Hide_all_buttons()
        btnNewProject.Visible = False
        grp_ProjectManager.Visible = False
        grp_MailManager.Visible = False
        grp_QualityManager.Visible = False
        grp_DrawingManager.Visible = False
        grp_TimeManager.Visible = False
        grp_TaskManager.Visible = False
        grp_ProjectGroupSettings.Visible = False
        grp_General.Visible = False
    End Sub

    Sub Reset_ribbon_from_Database(access As EffectiveAccessDto)

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        SharedRibbonButtons.CurrentAccess = access

        If access Is Nothing Then
            Hide_all_buttons()
            Cursor.Current = Cursors.Default
            Exit Sub
        End If

        SharedRibbonButtons.update_app_name()

        TabstructIQe.Label = If(String.IsNullOrWhiteSpace(access.Project_Assist_Name),
                            SettingsHelper.App_Name, access.Project_Assist_Name)

        If access.EffPaEnabled = False Then
            Hide_all_buttons()
            Cursor.Current = Cursors.Default
            Exit Sub
        End If

        grp_General.Visible = True
        grp_ProjectManager.Visible = True
        '.

        ' Feature-controlled items
        btnNewProject.Visible = SupabaseHelper.Has(access, "projects")
        grp_MailManager.Visible = SupabaseHelper.Has(access, "mail")
        grp_QualityManager.Visible = SupabaseHelper.Has(access, "quality")
        grp_DrawingManager.Visible = SupabaseHelper.Has(access, "drawings")
        grp_TimeManager.Visible = SupabaseHelper.Has(access, "timesheet")
        grp_TaskManager.Visible = SupabaseHelper.Has(access, "tasks")
        grp_ProjectGroupSettings.Visible = SupabaseHelper.Has(access, "project_group_manager")

        Cursor.Current = Cursors.Default

    End Sub
    Private Sub BtnNewProject_Click(sender As Object, e As RibbonControlEventArgs) Handles btnNewProject.Click

        SharedRibbonButtons.Button_New_Project(False)

    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click

        SharedRibbonButtons.Button_Manage_Projects()

    End Sub


    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click

        ThisAddIn.Button_File_mails()

    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click

        Shared_MailManagement_Class.Button_Retrieve_mails()

    End Sub


    Private Sub BtnSubmit_for_QC_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSubmit_for_QC.Click

        SharedQualityCheckClass.Button_Submit_for_QC("", "")

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click

        MsgBox("Under Development", MsgBoxStyle.Information, SettingsHelper.App_Name)

        ' SharedQualityCheckClass.Button_Submit_QC_Status()

    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click

        SharedQualityCheckClass.Button_Review_QC_Status()

    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLicenseOptions.Click


        If SupabaseHelper.IsOnline() = False Then
            Exit Sub
        End If

        Dim access

        Dim res = SharedRibbonButtons.User_Login()
        If res Is Nothing Then
            Exit Sub
        ElseIf res = True Then
            access = SharedRibbonButtons.CurrentAccess
        Else
            access = Nothing
        End If

        Reset_ribbon_from_Database(access)

    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click

        SharedRibbonButtons.Button_ProjectGroup_Edit_Settings()

    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click

        SharedRibbonButtons.Button_About_Box()

    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        SharedRibbonButtons.Button_User_Profiles()
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        SharedRibbonButtons.Button_Submit_Timesheet()
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        SharedRibbonButtons.Button_New_Project_Group()
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        MsgBox("Under Development", MsgBoxStyle.Information, SettingsHelper.App_Name)

    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        MsgBox("Under Development", MsgBoxStyle.Information, SettingsHelper.App_Name)
    End Sub

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        MsgBox("Under Development", MsgBoxStyle.Information, SettingsHelper.App_Name)
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        MsgBox("Under Development", MsgBoxStyle.Information, SettingsHelper.App_Name)
    End Sub

    Private Sub Button18_Click(sender As Object, e As RibbonControlEventArgs) Handles Button18.Click

        MsgBox("Under Development", MsgBoxStyle.Information, SettingsHelper.App_Name)

        'SharedRibbonButtons.Button_Manage_Drawings()
    End Sub

    Private Sub ButtonHelp_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonHelp.Click
        SharedRibbonButtons.Button_Help()
    End Sub

    Private Sub Button_Switch_Accounts_Click(sender As Object, e As RibbonControlEventArgs) Handles Button_Switch_Accounts.Click

        If SupabaseHelper.IsOnline() = False Then
            Exit Sub
        End If

        If SharedRibbonButtons.Button_Switch_User() Then
            Refresh_Ribbon_As_per_permissions()
        End If


    End Sub
End Class
