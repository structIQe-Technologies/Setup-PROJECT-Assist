Imports System.ComponentModel.Design
Imports System.Diagnostics
Imports System.Net
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Common_Library
Imports Microsoft.Office.Tools.Ribbon
Imports PROJECT_Assist_Common_Library
Imports structIQe_Common_Library

Public Class StructIQe

    Private DeferTimer As Windows.Forms.Timer
    Private StartupSw As New Stopwatch()

    Private Sub StructIQe_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        TabstructIQe.Label = Shared_Settings.App_Name

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

            Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        Catch ex As System.Exception
        Finally
            sw.Stop()
        End Try
    End Sub

    Async Function Refresh_Ribbon_As_per_permissions() As Task(Of Boolean)

        If SupabaseHelper.IsOnline() = False Then
            Hide_all_buttons()
            Return False
        End If

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        Dim access = Await SupabaseHelper.TryAutoLoginAsync()

        If access Is Nothing Then
            Hide_all_buttons()
            Dim res = SharedRibbonButtons.User_Login()
            If res Is Nothing Then
                System.Windows.Forms.Cursor.Current = Cursors.Default
                Return False
            ElseIf res = True Then
                access = Shared_Settings.CurrentAccess
                Shared_Settings.Update_app_name()
            Else
                access = Nothing
            End If

        End If

        Reset_ribbon_from_Database(access)

        System.Windows.Forms.Cursor.Current = Cursors.Default

        Return True

    End Function

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

        Shared_Settings.CurrentAccess = access

        If access Is Nothing Then
            Hide_all_buttons()
            Cursor.Current = Cursors.Default
            Exit Sub
        End If


        Dim application_name As String = Shared_Settings.Update_app_name()

        TabstructIQe.Label = If(String.IsNullOrWhiteSpace(application_name), Shared_Settings.App_Name, application_name)

        'If access.EffPaEnabled = False Then
        '    Hide_all_buttons()
        '    Cursor.Current = Cursors.Default
        '    Exit Sub
        'End If

        If Shared_Settings.IsSoftwareEnabled("project_assist") = False Then
            Hide_all_buttons()
            Cursor.Current = Cursors.Default
            Exit Sub
        End If

        grp_General.Visible = True
        grp_ProjectManager.Visible = True

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

        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        SharedRibbonButtons.Button_New_Project(False)

    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click

        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        SharedRibbonButtons.Button_Manage_Projects()

    End Sub


    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click

        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        ThisAddIn.Button_File_mails()

    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click

        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        SharedRibbonButtons.Button_Retrieve_mails()

    End Sub


    Private Sub BtnSubmit_for_QC_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSubmit_for_QC.Click

        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        SharedQualityCheckClass.Button_Submit_for_QC("", "")

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click

        MsgBox("Under Development", MsgBoxStyle.Information, Shared_Settings.App_Name)

        ' SharedQualityCheckClass.Button_Submit_QC_Status()

    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click

        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        SharedQualityCheckClass.Button_Review_QC_Status()

    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogin.Click


        If SupabaseHelper.IsOnline() = False Then
            Exit Sub
        End If

        Dim access

        Dim res = SharedRibbonButtons.User_Login()
        If res Is Nothing Then
            Exit Sub
        ElseIf res = True Then
            access = Shared_Settings.CurrentAccess
        Else
            access = Nothing
        End If

        Reset_ribbon_from_Database(access)

    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click

        'SecureStore.RestoreLastActiveSessionAndApplyToken(persist:=True)
        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        SharedRibbonButtons.Button_ProjectGroup_Edit_Settings()

    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click


        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        SharedRibbonButtons.Button_About_Box()

    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        'SecureStore.RestoreLastActiveSessionAndApplyToken(persist:=True)
        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()
        SharedRibbonButtons.Button_User_Profiles()
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()
        SharedRibbonButtons.Button_Submit_Timesheet()
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()
        SharedRibbonButtons.Button_New_Project_Group()
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        MsgBox("Under Development", MsgBoxStyle.Information, Shared_Settings.App_Name)

    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        MsgBox("Under Development", MsgBoxStyle.Information, Shared_Settings.App_Name)
    End Sub

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        MsgBox("Under Development", MsgBoxStyle.Information, Shared_Settings.App_Name)
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        MsgBox("Under Development", MsgBoxStyle.Information, Shared_Settings.App_Name)
    End Sub

    Private Sub Button18_Click(sender As Object, e As RibbonControlEventArgs) Handles Button18.Click

        MsgBox("Under Development", MsgBoxStyle.Information, Shared_Settings.App_Name)

        'SharedRibbonButtons.Button_Manage_Drawings()
    End Sub

    Private Sub ButtonHelp_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonHelp.Click
        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()
        SharedRibbonButtons.Button_Help()
    End Sub

    Private Async Sub Button_Switch_Accounts_Click(sender As Object, e As RibbonControlEventArgs) Handles Button_Switch_Accounts.Click

        If SupabaseHelper.IsOnline() = False Then
            Exit Sub
        End If

        If SharedRibbonButtons.Button_Switch_User() Then
            Await Refresh_Ribbon_As_per_permissions()
        End If

    End Sub

    Private Async Sub ButtonRefresh_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonRefresh.Click

        Refresh_Ribbon_As_per_permissions().GetAwaiter.GetResult()

        Await Common_Library.Cache_Builder.EnsureCache(Shared_Settings.CurrentAccess.CompanyId, forceReload:=True).ConfigureAwait(False)

    End Sub
End Class
