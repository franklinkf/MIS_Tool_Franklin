'Note: Codes related to Excel and Word file creation commented. Option25 and Option28 and functions for excel file creation.

Imports Oracle.DataAccess.Client
Imports System
Imports System.IO
Imports System.Data
Imports System.Diagnostics
Imports System.Reflection
Imports System.IO.Compression
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word

Public Class MISTool
    Dim rptoption As Integer = 0
    Dim datecheck As Boolean
    Dim RptDate As Date
    Dim tempvar As String
    Dim tempvar1 As String
    Dim tempvar2 As String
    Dim tempcount As Integer
    Dim tempdate As Date
    Dim tempdate1 As Date
    Dim tempdate2 As Date
    Dim Disk As String
    Dim folderpath As String = "c:\du"
    Dim file1 As String
    Dim file2 As String
    Dim file3 As String
    Dim file4 As String
    Dim file5 As String
    Dim file6 As String
    Dim sql As String
    Dim username As String

    Dim int_change_date_no_of_months As Integer
    Dim int_change_date_theo_balance As Double
    Dim int_change_date_inst_date As Date
    Dim overdueason As Date
    Dim txtnewod As Double
    Dim accountbalance As Double
    Dim dtintchangedate As Date
    Dim txtemi As Double
    Dim acno As String

    Dim cnn As New OleDb.OleDbConnection
    Dim cmd As OleDb.OleDbCommand
    Dim flag As Integer = 0
    Dim menulist(1000, 3) As String
    Dim menuitems_count = 80
    Dim gprocessid As Integer
    Dim gprocessname As String

    Private Sub Button1_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles Button1.Click
        ' Validating report option

        If txtusername.Text = "" Then

            MsgBox("Enter user name", MsgBoxStyle.Critical, "Enter User Name")
            Exit Sub

        Else

            username = txtusername.Text

        End If

        'CalculateEMI()

        If IsNumeric(txtcode.Text) = False Then

            MsgBox("Select valid option", MsgBoxStyle.Critical, "Invalid Option")
            Exit Sub

        Else

            If txtmenu.Text = "" Then
                MsgBox("Select valid option", MsgBoxStyle.Critical, "Invalid Option")
                Exit Sub
            Else
                rptoption = txtcode.Text
            End If

        End If

        If rptoption = 0 Then

            MsgBox("Select valid option", MsgBoxStyle.Critical, "Invalid Option")
            Exit Sub

        End If

        ' Validating date and assigning to variable

        Try

            RptDate = CDate(txtdate.Text)

        Catch ex As Exception

            MsgBox("Enter valid date", MsgBoxStyle.Critical, "Invalid date")
            Exit Sub

        End Try

        ' Checking folder path, if do not exists, create the same.

        processmessage("Checking folder path")

        If Directory.Exists(folderpath) Then

            tempvar = "aa"

        Else

            Directory.CreateDirectory("c:\du")

        End If

        If rptoption = 1 Then

            option1()

        ElseIf rptoption = 2 Then

            option2()

        ElseIf rptoption = 3 Then

            option3()

        ElseIf rptoption = 4 Then

            option4()

        ElseIf rptoption = 5 Then

            option5()

        ElseIf rptoption = 6 Then

            option6()

        ElseIf rptoption = 7 Then

            option7()

        ElseIf rptoption = 8 Then

            option8()

        ElseIf rptoption = 9 Then

            option9()

        ElseIf rptoption = 10 Then

            Button3.Enabled = True

            option10()

        ElseIf rptoption = 11 Then

            Button3.Enabled = True

            option11()

        ElseIf rptoption = 12 Then

            Button3.Enabled = True

            option12()

        ElseIf rptoption = 13 Then

            option13()

        ElseIf rptoption = 14 Then

            option14()

        ElseIf rptoption = 15 Then

            option15()

        ElseIf rptoption = 16 Then

            option16()

        ElseIf rptoption = 17 Then

            option17()

        ElseIf rptoption = 18 Then

            option18()

        ElseIf rptoption = 19 Then

            option19()

        ElseIf rptoption = 20 Then

            option20()

        ElseIf rptoption = 21 Then

            option21()

        ElseIf rptoption = 9072 Then

            option9072()

        ElseIf rptoption = 9074 Then

            option9074()

        ElseIf rptoption = 9071 Then

            option9071()

            'ElseIf rptoption = 25 Then

            '    option25()

        ElseIf rptoption = 26 Then

            option26()

        ElseIf rptoption = 27 Then

            option27()

            'ElseIf rptoption = 28 Then

            '    option28()

        ElseIf rptoption = 29 Then

            option29()

        ElseIf rptoption = 30 Then

            option30()

        ElseIf rptoption = 31 Then

            option31()

        ElseIf rptoption = 32 Then

            option32()

        ElseIf rptoption = 33 Then

            option33()

        ElseIf rptoption = 34 Then

            option34()

        ElseIf rptoption = 35 Then

            option35()

        ElseIf rptoption = 36 Then

            option36()

        ElseIf rptoption = 37 Then

            option37()

        ElseIf rptoption = 38 Then

            option38()

        ElseIf rptoption = 39 Then

            option39()

        ElseIf rptoption = 40 Then

            option40()

        ElseIf rptoption = 41 Then

            option41()

        ElseIf rptoption = 42 Then

            option42()

        ElseIf rptoption = 43 Then

            option43()

        ElseIf rptoption = 44 Then

            option44()

        ElseIf rptoption = 45 Then

            option45()

        ElseIf rptoption = 46 Then

            option46()

        ElseIf rptoption = 47 Then

            option47()

        ElseIf rptoption = 48 Then

            option48()

        ElseIf rptoption = 53 Then

            option53()

        ElseIf rptoption = 54 Then

            option54()

        ElseIf rptoption = 55 Then

            Option55()

        ElseIf rptoption = 56 Then

            option56()

        ElseIf rptoption = 57 Then

            Option57()

        ElseIf rptoption = 58 Then

            option58()

        ElseIf rptoption = 59 Then

            option59()

        ElseIf rptoption = 60 Then

            option60()

        ElseIf rptoption = 61 Then

            option61()

        ElseIf rptoption = 62 Then

            option62()
        ElseIf rptoption = 63 Then

            option63()

        ElseIf rptoption = 64 Then

            option64()

        ElseIf rptoption = 65 Then

            option65()
        ElseIf rptoption = 66 Then

            option66()

        ElseIf rptoption = 67 Then

            option67()
        ElseIf rptoption = 68 Then

            option68()


        ElseIf rptoption = 801 Then

            option801()

        ElseIf rptoption = 802 Then

            option802()

        ElseIf rptoption = 803 Then

            option803()

        ElseIf rptoption = 804 Then

            option804()

        ElseIf rptoption = 805 Then

            option805()

        ElseIf rptoption = 806 Then

            option806()

        ElseIf rptoption = 807 Then

            option807()

        ElseIf rptoption = 808 Then

            option808()

        ElseIf rptoption = 809 Then

            option809()

        ElseIf rptoption = 810 Then

            option810()

        ElseIf rptoption = 811 Then

            option811()

        ElseIf rptoption = 812 Then

            option812()

        ElseIf rptoption = 813 Then

            option813()

        ElseIf rptoption = 814 Then

            option814()

        ElseIf rptoption = 815 Then

            option815()

        ElseIf rptoption = 816 Then

            option816()

        ElseIf rptoption = 817 Then

            option817()

        ElseIf rptoption = 818 Then

            option818()

        ElseIf rptoption = 819 Then

            option819()

        ElseIf rptoption = 820 Then

            option820()

        ElseIf rptoption = 821 Then

            option821()

        ElseIf rptoption = 822 Then

            option822()

        ElseIf rptoption = 823 Then

            option823()

        ElseIf rptoption = 824 Then

            option824()

        ElseIf rptoption = 825 Then

            option825()

        ElseIf rptoption = 826 Then

            Option826()

        ElseIf rptoption = 827 Then

            option827()

        ElseIf rptoption = 828 Then

            option828()

        ElseIf rptoption = 829 Then

            Option829()

        ElseIf rptoption = 830 Then

            Option830()

        ElseIf rptoption = 831 Then

            option831()

        ElseIf rptoption = 832 Then

            option832()

        ElseIf rptoption = 833 Then

            option833()

        ElseIf rptoption = 601 Then

            option601()

        ElseIf rptoption = 604 Then

            option604()

        ElseIf rptoption = 603 Then

            option603()

        ElseIf rptoption = 602 Then

            option602()

        ElseIf rptoption = 605 Then

            option605()

        ElseIf rptoption = 606 Then

            option606()

        ElseIf rptoption = 607 Then

            option607()

        ElseIf rptoption = 608 Then

            option608()

        ElseIf rptoption = 609 Then

            option609()

        ElseIf rptoption = 610 Then

            option610()

        ElseIf rptoption = 611 Then

            option611()

        ElseIf rptoption = 612 Then

            option612()

        ElseIf rptoption = 613 Then

            option613()

        ElseIf rptoption = 614 Then

            option614()

        ElseIf rptoption = 615 Then

            option615()

        ElseIf rptoption = 616 Then

            option616()

        ElseIf rptoption = 617 Then

            option617()

        End If

        '' FRANKLIN - DEFINE MENULIST NEXT (SEARCH FOR >> FRANKLIN - DEFINE MENU LIST)

        clearall()
        txtcode.Text = ""
        txtcode.Focus()

    End Sub


    Public Sub display_in_File(ByVal sql, ByVal filename) 'Display the output of the sql in filename

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim cmd1 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd1.ExecuteReader()
        Dim linedata As String
        Dim sw As StreamWriter = New StreamWriter(filename.ToString())
        While dr.Read()
            linedata = ""
            linedata = dr("REPORTDATA").ToString()
            sw.WriteLine(linedata)
        End While
        dr.Close()
        sw.Close()
        cnn.Close()
    End Sub



    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Me.Text = "MIS Tool (V 20)"
        Button3.Enabled = False
        txtdate.Text = Today.AddDays(-1)
        lblstatus.Text = ""
        lblstatus2.Text = ""
        lblinfo1.Text = ""
        lblinfo2.Text = ""
        lblinfo3.Text = ""
        lblinfo4.Text = ""
        lblinfo5.Text = "Select Option"
        lblinfo6.Text = ""
        lblinfo7.Text = ""
        lblinfo8.Text = ""
        lblinfo9.Text = ""
        lblinfo10.Text = ""
        dgv1.Visible = False
        txtcode.Focus()
        Dim sr As StreamReader = New StreamReader("C:/MT.txt")
        Do While sr.Peek() >= 0
            tempvar = sr.ReadLine()
            Disk = tempvar.Substring(0, 1)
        Loop
        If Disk = "" Then
            Disk = "E"
        End If

        menulist(0, 0) = "1"
        menulist(0, 1) = "Aadhaar Upload - Delete Duplicate Records"
        menulist(0, 2) = "A"

        menulist(1, 0) = "2"
        menulist(1, 1) = "Daily emails"
        menulist(1, 2) = "A"

        menulist(2, 0) = "3"
        menulist(2, 1) = "Upload Files"
        menulist(2, 2) = "A"

        menulist(3, 0) = "4"
        menulist(3, 1) = "Tabdata"
        menulist(3, 2) = "A"

        menulist(4, 0) = "5"
        menulist(4, 1) = "General"
        menulist(4, 2) = "A"

        menulist(5, 0) = "6"
        menulist(5, 1) = "Report"
        menulist(5, 2) = "A"

        menulist(6, 0) = "7"
        menulist(6, 1) = "KGB Business Progress Report"
        menulist(6, 2) = "A"

        menulist(7, 0) = "8"
        menulist(7, 1) = "KGB Day Book"
        menulist(7, 2) = "A"

        menulist(8, 0) = "9"
        menulist(8, 1) = "Business Review"
        menulist(8, 2) = "A"

        menulist(9, 0) = "10"
        menulist(9, 1) = "KGB First - Outstanding"
        menulist(9, 2) = "A"

        menulist(10, 0) = "11"
        menulist(10, 1) = "KGB First - Disbursement"
        menulist(10, 2) = "A"
        menulist(11, 0) = "12"
        menulist(11, 1) = "KGB First - NPA"
        menulist(11, 2) = "A"
        menulist(12, 0) = "13"
        menulist(12, 1) = "MISDO Upload"
        menulist(12, 2) = "A"
        menulist(13, 0) = "14"
        menulist(13, 1) = "ATM Data Mismatch between Finacle and Switch reports"
        menulist(13, 2) = "A"
        menulist(14, 0) = "15"
        menulist(14, 1) = "CIBIL Upload File Creation (Live)"
        menulist(14, 2) = "A"
        menulist(15, 0) = "16"
        menulist(15, 1) = "EMail Daily Reports"
        menulist(15, 2) = "A"
        menulist(16, 0) = "17"
        menulist(16, 1) = "NPCI Linked Aadhaar - Upload file creation"
        menulist(16, 2) = "A"
        menulist(17, 0) = "18"
        menulist(17, 1) = "Day end eMails"
        menulist(17, 2) = "A"
        menulist(18, 0) = "19"
        menulist(18, 1) = "Business Review - Files to RO"
        menulist(18, 2) = "A"
        menulist(19, 0) = "20"
        menulist(19, 1) = "KGB Aadhar Enrolled Status"
        menulist(19, 2) = "A"
       
        menulist(20, 0) = "21"
        menulist(20, 1) = "KGB Daily Reports"
        menulist(20, 2) = "A"
        menulist(21, 0) = "9072"            '  "22" OPTION= 22 CHANGED AS 9072 FOR READABILITY
        menulist(21, 1) = "9072 Insert"
        menulist(21, 2) = "A"
        menulist(22, 0) = "9074"            '  "23" OPTION= 23 CHANGED AS 9074 FOR READABILITY
        menulist(22, 1) = "9074 Insert"
        menulist(22, 2) = "A"
        menulist(23, 0) = "9071"              '  "24" OPTION= 24 CHANGED AS 9071 FOR READABILITY
        menulist(23, 1) = "9071 Insert"
        menulist(23, 2) = "A"
        menulist(24, 0) = "25"
        menulist(24, 1) = "Create RO and Branch Folders and convert CIB Files"
        menulist(24, 2) = "A"
        menulist(25, 0) = "26"
        menulist(25, 1) = "Create Bank as a whole/All RO's/All Branches report in a single file"
        menulist(25, 2) = "A"
        menulist(26, 0) = "27"
        menulist(26, 1) = "Get File Names"
        menulist(26, 2) = "A"
        menulist(27, 0) = "28"
        menulist(27, 1) = "Word Document Generation"
        menulist(27, 2) = "A"
        menulist(28, 0) = "29"
        menulist(28, 1) = "Mobile Banking Transaction Status"
        menulist(28, 2) = "A"
        menulist(29, 0) = "30"
        menulist(29, 1) = "Create Folder"
        menulist(29, 2) = "A"
       
        menulist(30, 0) = "31"
        menulist(30, 1) = "Copy File"
        menulist(30, 2) = "A"
        menulist(31, 0) = "32"
        menulist(31, 1) = "Execute Script"
        menulist(31, 2) = "A"
        menulist(32, 0) = "33"
        menulist(32, 1) = "Basedata Generation Timing"
        menulist(32, 2) = "A"
        menulist(33, 0) = "34"
        menulist(33, 1) = "Staff Upload"
        menulist(33, 2) = "A"
        menulist(34, 0) = "35"
        menulist(34, 1) = "RO Follow Up Status Email"
        menulist(34, 2) = "A"
        menulist(35, 0) = "36"
        menulist(35, 1) = "ATM Transaction Status"
        menulist(35, 2) = "A"
        menulist(36, 0) = "37"
        menulist(36, 1) = "Upload data to tables - All Columns"
        menulist(36, 2) = "A"
        menulist(37, 0) = "38"
        menulist(37, 1) = "Upload data to tables - Partial Columns"
        menulist(37, 2) = "A"
        menulist(38, 0) = "39"
        menulist(38, 1) = "Migration Tool Data Entry Status Email"
        menulist(38, 2) = "A"
        menulist(39, 0) = "40"
        menulist(39, 1) = "Export Oracle Data"
        menulist(39, 2) = "A"
        
        menulist(40, 0) = "41"
        menulist(40, 1) = "Backup, Drop and Import Oracle Tables"
        menulist(40, 2) = "A"
        menulist(41, 0) = "601"
        menulist(41, 1) = "eNMGB Migration - Create Branch Data"
        menulist(41, 2) = "A"
        menulist(42, 0) = "602"
        menulist(42, 1) = "eNMGB Migration - Upload Migration Tool Files"
        menulist(42, 2) = "A"
        menulist(43, 0) = "603"
        menulist(43, 1) = "eNMGB Migration - Upload CGL File"
        menulist(43, 2) = "A"
        menulist(44, 0) = "604"
        menulist(44, 1) = "eNMGB Migration - Assign CustID and Account No"
        menulist(44, 2) = "A"
        menulist(45, 0) = "605"
        menulist(45, 1) = "eNMGB Migration - FUF Generation"
        menulist(45, 2) = "A"
        menulist(46, 0) = "606"
        menulist(46, 1) = "eNMGB Migration - Reports"
        menulist(46, 2) = "A"
        menulist(47, 0) = "607"
        menulist(47, 1) = "eNMGB Migration - Upload 2059 Files"
        menulist(47, 2) = "A"
        menulist(48, 0) = "608"
        menulist(48, 1) = "eNMGB Migration - Check 2059 Files"
        menulist(48, 2) = "A"
        menulist(49, 0) = "609"
        menulist(49, 1) = "eNMGB Migration - Split CEDGE Dump"
        menulist(49, 2) = "A"
        menulist(50, 0) = "612"
        menulist(50, 1) = "eNMGB Migration - Batch update of packages"
        menulist(50, 2) = "A"
        menulist(51, 0) = "610"
        menulist(51, 1) = "eNMGB Migration - Create History Transaction Data Dump"
        menulist(51, 2) = "A"
        menulist(52, 0) = "611"
        menulist(52, 1) = "eNMGB Migration - Create NPA Upload Files"
        menulist(52, 2) = "A"
        menulist(53, 0) = "613"
        menulist(53, 1) = "eNMGB Migration - Create backup of live users"
        menulist(53, 2) = "A"
        menulist(54, 0) = "614"
        menulist(54, 1) = "eNMGB Migration - Import Users"
        menulist(54, 2) = "A"
        menulist(55, 0) = "615"
        menulist(55, 1) = "eNMGB Migration - Data from users"
        menulist(55, 2) = "A"
        menulist(56, 0) = "42"
        menulist(56, 1) = "Drop oracle user"
        menulist(56, 2) = "A"
        menulist(57, 0) = "43"
        menulist(57, 1) = "Figures At A Glance"
        menulist(57, 2) = "A"
        menulist(58, 0) = "616"
        menulist(58, 1) = "eNMGB Migration - Zenith Backup Import"
        menulist(58, 2) = "A"
        menulist(59, 0) = "44"
        menulist(59, 1) = "Execute query and generate multiple files"
        menulist(59, 2) = "A"
        menulist(60, 0) = "45"
        menulist(60, 1) = "PMJDY Campaign"
        menulist(60, 2) = "A"
        menulist(61, 0) = "46"
        menulist(61, 1) = "Bulk SMS File Creation"
        menulist(61, 2) = "A"
        menulist(62, 0) = "47"
        menulist(62, 1) = "Business Figures As On 30-09-2014"
        menulist(62, 2) = "A"
        menulist(63, 0) = "48"
        menulist(63, 1) = "Branch Intimation Letter"
        menulist(63, 2) = "A"
        'menulist(64, 0) = "53"
        'menulist(64, 1) = "SARFAESI Notice Intimation Status"
        'menulist(64, 2) = "A"
        menulist(64, 0) = "53"
        menulist(64, 1) = "PMJJBY/PMSBY/APY Enrollment Status"
        menulist(64, 2) = "A"
        menulist(65, 0) = "55"
        menulist(65, 1) = "NPA Threat For Next 7 Days - Email Generation"
        menulist(65, 2) = "A"
        menulist(66, 0) = "54"
        menulist(66, 1) = "NPA Threat For Next 7 Days - Excel Creation"
        menulist(66, 2) = "A"
        menulist(67, 0) = "56"
        menulist(67, 1) = "NPA Threat For Next 7 Days - Excel Creation Using Macro"
        menulist(67, 2) = "A"
        menulist(68, 0) = "57"
        menulist(68, 1) = "Predefined Day End Check Validation"
        menulist(68, 2) = "A"
        menulist(69, 0) = "58"
        menulist(69, 1) = "BOD Mails"
        menulist(69, 2) = "A"

        menulist(70, 0) = "59"
        menulist(70, 1) = "NPA Reports"
        menulist(70, 2) = "A"
        menulist(71, 0) = "60"
        menulist(71, 1) = "PROGRESS REPORT AS PER CIRCULAR: 74/2015"
        menulist(71, 2) = "A"
        menulist(72, 0) = "61"
        menulist(72, 1) = "KYC Upload Statistics"
        menulist(72, 2) = "A"

        menulist(73, 0) = "62"
        menulist(73, 1) = "MASS NEFT AGRICULTURE DEPT"
        menulist(73, 2) = "A"

        menulist(74, 0) = "63"
        menulist(74, 1) = "Kiosk file"
        menulist(74, 2) = "A"

        menulist(75, 0) = "64"
        menulist(75, 1) = "Weekly Transaction Mail"
        menulist(75, 2) = "A"

        menulist(76, 0) = "65"
        menulist(76, 1) = "Data Upload for DashBoard"
        menulist(76, 2) = "A"

        menulist(77, 0) = "66"
        menulist(77, 1) = "CIBIL Upload File Creation (Close)"
        menulist(77, 2) = "A"

        menulist(78, 0) = "67"
        menulist(78, 1) = "Mobile banking SMS creation"
        menulist(78, 2) = "A"

        menulist(79, 0) = "68"
        menulist(79, 1) = "Transaction Data upload for DashBoard"
        menulist(79, 2) = "A"

        menulist(80, 0) = "617"
        menulist(80, 1) = "20 Twenty Session Batch Job"
        menulist(80, 2) = "A"

        '' Add one more entry in  Dim menuitems_count = XX
        '' FRANKLIN - DEFINE MENU LIST
        '' NEXT >> UPDATE LABEL INFO >> SEARCH FOR >> FRANKLIN - UPDATE LABEL INFO

        '----------------------------------------------------------------------------------------------------------
        'CODE FOR INSERTING DATA TO MIGRATION TOOL -- ACCESS DATABASE NMGB.MDB
        '----------------------------------------------------------------------------------------------------------
        'menulist(50, 0) = "801"
        'menulist(50, 1) = "Inserting data into Location table"
        'menulist(50, 2) = "A"
        'menulist(51, 0) = "802"
        'menulist(51, 1) = "Inserting data into CIDMASTER table"
        'menulist(51, 2) = "A"
        'menulist(52, 0) = "803"
        'menulist(52, 1) = "Inserting data to Pickup table"
        'menulist(52, 2) = "A"
        'menulist(53, 0) = "804"
        'menulist(53, 1) = "Inserting data to Religioncode table"
        'menulist(53, 2) = "A"
        'menulist(54, 0) = "805"
        'menulist(54, 1) = "Update religioncode from banc724"
        'menulist(54, 2) = "A"
        'menulist(55, 0) = "806"
        'menulist(55, 1) = "Inserting data to BranchMaster"
        'menulist(55, 2) = "A"
        'menulist(56, 0) = "807"
        'menulist(56, 1) = "Inserting Deposit shadow file"
        'menulist(56, 2) = "A"
        'menulist(57, 0) = "808"
        'menulist(57, 1) = "Inserting Loan shadow file"
        'menulist(57, 2) = "A"
        'menulist(58, 0) = "809"
        'menulist(58, 1) = "Updating NRE code"
        'menulist(58, 2) = "A"
        'menulist(59, 0) = "810"
        'menulist(59, 1) = "Inserting Staff Code"
        'menulist(59, 2) = "A"

        'menulist(60, 0) = "811"
        'menulist(60, 1) = "Category code"
        'menulist(60, 2) = "A"
        'menulist(61, 0) = "812"
        'menulist(61, 1) = "Inserting data to Citycode1"
        'menulist(61, 2) = "A"
        'menulist(62, 0) = "813"
        'menulist(62, 1) = "Inserting data to Citycode2"
        'menulist(62, 2) = "A"
        'menulist(63, 0) = "814"
        'menulist(63, 1) = "Inserting data to Minor table"
        'menulist(63, 2) = "A"
        'menulist(64, 0) = "815"
        'menulist(64, 1) = "uncompress"
        'menulist(64, 2) = "A"
        'menulist(65, 0) = "816"
        'menulist(65, 1) = "Inserting Param file and database"
        'menulist(65, 2) = "A"
        'menulist(66, 0) = "817"
        'menulist(66, 1) = "Copying files for Creating Setup"
        'menulist(66, 2) = "A"
        'menulist(67, 0) = "818"
        'menulist(67, 1) = "NRE from file"
        'menulist(67, 2) = "A"
        'menulist(68, 0) = "819"
        'menulist(68, 1) = "Deceased from file"
        'menulist(68, 2) = "A"
        'menulist(69, 0) = "820"
        'menulist(69, 1) = "Staff no From file"
        'menulist(69, 2) = "A"

        'menulist(70, 0) = "821"
        'menulist(70, 1) = "Category from file"
        'menulist(70, 2) = "A"
        'menulist(71, 0) = "822"
        'menulist(71, 1) = "Religion from file"
        'menulist(71, 2) = "A"
        'menulist(72, 0) = "823"
        'menulist(72, 1) = "Handicapped from file"
        'menulist(72, 2) = "A"
        'menulist(73, 0) = "824"
        'menulist(73, 1) = "LPD details from file"
        'menulist(73, 2) = "A"
        'menulist(74, 0) = "825"
        'menulist(74, 1) = "Compress and email"
        'menulist(74, 2) = "A"
        'menulist(75, 0) = "826"
        'menulist(75, 1) = "Differential Backup"
        'menulist(75, 2) = "A"
        'menulist(76, 0) = "827"
        'menulist(76, 1) = "Upload - Extension based"
        'menulist(76, 2) = "A"
        'menulist(77, 0) = "828"
        'menulist(77, 1) = "Insert into tables"
        'menulist(77, 2) = "A"
        'menulist(78, 0) = "829"
        'menulist(78, 1) = "Differential Backup based on Extension"
        'menulist(78, 2) = "A"
        'menulist(79, 0) = "830"
        'menulist(79, 1) = "Mirror image"
        'menulist(79, 2) = "A"

        'menulist(80, 0) = "831"
        'menulist(80, 1) = "Generating CIDMaster File From dump"
        'menulist(80, 2) = "A"
        'menulist(81, 0) = "832"
        'menulist(81, 1) = "Create text files in a loop"
        'menulist(81, 2) = "A"
        'menulist(82, 0) = "833"
        'menulist(82, 1) = "Citycode 3 -- Issue"
        'menulist(82, 2) = "A"

        '-----------------------------------------------------------------------------------------------

    End Sub

    Shared Function GetAccountForEmailAddress(ByVal application As Outlook.Application, ByVal smtpAddress As String) As Outlook.Account

        ' Loop over the Accounts collection of the current Outlook session.
        Dim accounts As Outlook.Accounts = application.Session.Accounts
        Dim account As Outlook.Account
        For Each account In accounts
            ' When the e-mail address matches, return the account.
            ' smtpAddress = "smgbmis3@gmail.com"
            If account.SmtpAddress = smtpAddress Then
                Return account
            End If
        Next
        Throw New System.Exception(String.Format("No Account with SmtpAddress: {0} exists!", smtpAddress))
    End Function

    Sub option1()       'Aadhaar Upload - Delete Duplicate Records

        'Checking whether Original.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\ORIGINAL.TXT"
        file2 = "c:\du\ERROR.TXT"

        checkfile(file1, "Aadhaar File not found in c:/du folder named as 'Original.txt'")
        checkfile(file2, "Error File not found in c:/du folder named as 'Error.txt'")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")

        ' Deleting erraneous records

        processmessage("Deleting erraneous records")

        oracle_execute_non_query("ten", username, username, "DELETE FROM Z_DU WHERE UPPER(FILENAME) = 'C:\DU\ORIGINAL.TXT' AND LINENO IN (SELECT TO_NUMBER(TRIM(SUBSTR(LINEDATA,INSTR(LINEDATA,'Line No: ')+9,10))) LINENO FROM Z_DU WHERE UPPER(FILENAME) = 'C:\DU\ERROR.TXT' AND INSTR(LINEDATA,'Line No: ') > 0)")

        processmessage("Creating new file")

        ' Creating new file

        Dim file3 As String = "c:\du\Aadhaar_New_File.txt"

        If File.Exists(file3) Then

            File.Delete(file3)

        End If

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim sw As StreamWriter = New StreamWriter(file3)
        sql = "SELECT LINEDATA FROM Z_DU WHERE UPPER(FILENAME) = 'C:\DU\ORIGINAL.TXT' ORDER BY LINENO"
        Dim cmd4 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd4.ExecuteReader()
        While dr.Read()
            tempvar = dr.Item("LINEDATA").ToString
            If tempvar <> "" Then
                sw.WriteLine(tempvar)
            End If
        End While
        dr.Close()
        sw.Close()

        processmessage("")

        MsgBox("New file named 'Aadhaar_New_File.txt' created in c:/du folder", MsgBoxStyle.Information, "Process Completed")

        ' Closing Oracle Connection

        conn.Close()
        conn.Dispose()

    End Sub

    Sub option2()       'Daily EMails

        ' Checking whether email.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\email.txt"
        'file2 = "c:\du\friday.txt"
        'file3 = "c:\du\kgbdb.txt"

        checkfile(file1, "Rename the EMail file 40101.email as email.txt and place in c:/du folder")
        'checkfile(file2, "Place last friday file in C:/DU folder as 'FRIDAY.txt'")
        'checkfile(file3, "Rename the MISDO file 40124.misdo as kgbdb.txt and place in c:/du folder")

        processmessage("Validating date")
        tempvar1 = readNthLine(file1, 0)

        Try

            tempdate1 = CDate(tempvar1)

        Catch ex As Exception

            MsgBox("Invalid date in first line of email.txt file", MsgBoxStyle.Critical, "Invalid date")
            Exit Sub

        End Try

        If RptDate <> tempdate1 Then

            MsgBox("Report date and date in email.txt do not match", MsgBoxStyle.Critical, "Mismatch in date")
            Exit Sub

        End If

        uploadfiledata(file1, username, "Y")
        'uploadfiledata(file2, username, "N")
        'uploadfiledata(file3, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - Get Data")

        sql = "PKGEMAIL102.GETDATA"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1012")      'NPA In Out

        'sql = "PKGEMAIL101.DATAID_1012"
        'Dim cmd7 As New OracleCommand(sql, conn)
        'cmd7.CommandType = CommandType.StoredProcedure
        'cmd7.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd7.ExecuteNonQuery()

        'sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'NPAINOUT' ORDER BY SOLID"
        'display_in_File(sql, "C:\du\SMS_" & RptDate.ToString("ddMMyyyy") & ".txt")

        'processmessage("Package - Data ID - 1013")      'Loans Opened

        'sql = "PKGEMAIL101.DATAID_1013"
        'Dim cmd8 As New OracleCommand(sql, conn)
        'cmd8.CommandType = CommandType.StoredProcedure
        'cmd8.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd8.ExecuteNonQuery()

        processmessage("Package - Data ID - 1021")      'Deposit

        sql = "PKGEMAIL102.DATAID_1021"
        Dim cmd10 As New OracleCommand(sql, conn)
        cmd10.CommandType = CommandType.StoredProcedure
        cmd10.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd10.ExecuteNonQuery()

        processmessage("Package - Data ID - 1022")      'Advance

        sql = "PKGEMAIL102.DATAID_1022"
        Dim cmd11 As New OracleCommand(sql, conn)
        cmd11.CommandType = CommandType.StoredProcedure
        cmd11.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd11.ExecuteNonQuery()

        processmessage("Package - Data ID - 1023")      'Aadhar Enrolment

        sql = "PKGEMAIL102.DATAID_1023"
        Dim cmd12 As New OracleCommand(sql, conn)
        cmd12.CommandType = CommandType.StoredProcedure
        cmd12.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd12.ExecuteNonQuery()

        processmessage("Package - Data ID - 1024")      'VBS Additional Data

        sql = "PKGEMAIL102.DATAID_1024"
        Dim cmd13 As New OracleCommand(sql, conn)
        cmd13.CommandType = CommandType.StoredProcedure
        cmd13.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd13.ExecuteNonQuery()

        processmessage("Package - Data ID - 1031")      'SMS Enrolment

        sql = "PKGEMAIL103.DATAID_1031"
        Dim cmd14 As New OracleCommand(sql, conn)
        cmd14.CommandType = CommandType.StoredProcedure
        cmd14.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd14.ExecuteNonQuery()

        processmessage("Package - Data ID - 1032")      'KYC 

        sql = "PKGEMAIL103.DATAID_1032"
        Dim cmd15 As New OracleCommand(sql, conn)
        cmd15.CommandType = CommandType.StoredProcedure
        cmd15.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd15.ExecuteNonQuery()

        processmessage("Package - Data ID - 1033")      'ATM Enrolment

        sql = "PKGEMAIL103.DATAID_1033"
        Dim cmd16 As New OracleCommand(sql, conn)
        cmd16.CommandType = CommandType.StoredProcedure
        cmd16.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd16.ExecuteNonQuery()

        processmessage("Package - Data ID - 1034")      'AOD Pending

        sql = "PKGEMAIL103.DATAID_1034"
        Dim cmd17 As New OracleCommand(sql, conn)
        cmd17.CommandType = CommandType.StoredProcedure
        cmd17.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd17.ExecuteNonQuery()

        processmessage("Package - Data ID - 1035")      'Location Updation

        sql = "PKGEMAIL103.DATAID_1035"
        Dim cmd18 As New OracleCommand(sql, conn)
        cmd18.CommandType = CommandType.StoredProcedure
        cmd18.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd18.ExecuteNonQuery()

        processmessage("Package - Data ID - 1041")      'Locker Status

        sql = "PKGEMAIL104.DATAID_1041"
        Dim cmd19 As New OracleCommand(sql, conn)
        cmd19.CommandType = CommandType.StoredProcedure
        cmd19.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd19.ExecuteNonQuery()

        processmessage("Package - Data ID - 1042")      'ABPS Remittance

        sql = "PKGEMAIL104.DATAID_1042"
        Dim cmd20 As New OracleCommand(sql, conn)
        cmd20.CommandType = CommandType.StoredProcedure
        cmd20.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd20.ExecuteNonQuery()

        processmessage("Package - Data ID - 1073")      'CIBIL Data Rectification

        sql = "PKGEMAIL107.DATAID_1073"
        Dim cmd30 As New OracleCommand(sql, conn)
        cmd30.CommandType = CommandType.StoredProcedure
        cmd30.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd30.ExecuteNonQuery()

        processmessage("Package - Data ID - 1074")      'LPD Module Data Entry

        sql = "PKGEMAIL107.DATAID_1074"
        Dim cmd31 As New OracleCommand(sql, conn)
        cmd31.CommandType = CommandType.StoredProcedure
        cmd31.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd31.ExecuteNonQuery()

        processmessage("Package - Data ID - 1076")      'CANBANK RRB NPA Data

        sql = "PKGEMAIL107.DATAID_1076"
        Dim cmd53 As New OracleCommand(sql, conn)
        cmd53.CommandType = CommandType.StoredProcedure
        cmd53.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd53.ExecuteNonQuery()

        processmessage("Package - Data ID - 1081")      'CERSAI Enrolment

        sql = "PKGEMAIL108.DATAID_1081"
        Dim cmd54 As New OracleCommand(sql, conn)
        cmd54.CommandType = CommandType.StoredProcedure
        cmd54.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd54.ExecuteNonQuery()

        processmessage("Package - Data ID - 1082")      'SL/SA/BAR/BILLS Pending

        sql = "PKGEMAIL108.DATAID_1082"
        Dim cmd55 As New OracleCommand(sql, conn)
        cmd55.CommandType = CommandType.StoredProcedure
        cmd55.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd55.ExecuteNonQuery()

        processmessage("Package - Data ID - 1083")      'ABPS Rejected Data

        sql = "PKGEMAIL108.DATAID_1083"
        Dim cmd56 As New OracleCommand(sql, conn)
        cmd56.CommandType = CommandType.StoredProcedure
        cmd56.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd56.ExecuteNonQuery()

        processmessage("Package - Data ID - 1084")      'Advances < 8%

        sql = "PKGEMAIL108.DATAID_1084"
        Dim cmd57 As New OracleCommand(sql, conn)
        cmd57.CommandType = CommandType.StoredProcedure
        cmd57.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd57.ExecuteNonQuery()

        processmessage("Package - Data ID - 1085")      'Deposits Pref Rate

        sql = "PKGEMAIL108.DATAID_1085"
        Dim cmd58 As New OracleCommand(sql, conn)
        cmd58.CommandType = CommandType.StoredProcedure
        cmd58.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd58.ExecuteNonQuery()

        processmessage("Package - Data ID - 1091")      'Daily Cash Position

        sql = "PKGEMAIL109.DATAID_1091"
        Dim cmd59 As New OracleCommand(sql, conn)
        cmd59.CommandType = CommandType.StoredProcedure
        cmd59.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd59.ExecuteNonQuery()

        processmessage("Package - Data ID - 1092")      'DBT Scheme Wise

        sql = "PKGEMAIL109.DATAID_1092"
        Dim cmd60 As New OracleCommand(sql, conn)
        cmd60.CommandType = CommandType.StoredProcedure
        cmd60.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd60.ExecuteNonQuery()

        processmessage("Package - Data ID - 1093")      'EM Not Entered In SRM

        sql = "PKGEMAIL109.DATAID_1093"
        Dim cmd61 As New OracleCommand(sql, conn)
        cmd61.CommandType = CommandType.StoredProcedure
        cmd61.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd61.ExecuteNonQuery()

        processmessage("Package - Data ID - 1094")      'A/c With Invalid SI Flag

        sql = "PKGEMAIL109.DATAID_1094"
        Dim cmd62 As New OracleCommand(sql, conn)
        cmd62.CommandType = CommandType.StoredProcedure
        cmd62.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd62.ExecuteNonQuery()

        processmessage("Package - Data ID - 1095")      'Loan A/c With Inadequate Security

        sql = "PKGEMAIL109.DATAID_1095"
        Dim cmd63 As New OracleCommand(sql, conn)
        cmd63.CommandType = CommandType.StoredProcedure
        cmd63.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd63.ExecuteNonQuery()

        processmessage("Package - Data ID - 1096")      'Issues in VSL Accounts

        sql = "PKGEMAIL109.DATAID_1096"
        Dim cmd64 As New OracleCommand(sql, conn)
        cmd64.CommandType = CommandType.StoredProcedure
        cmd64.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd64.ExecuteNonQuery()

        processmessage("Package - Data ID - 1101")      'Educational Loan Pending For Reschedule

        sql = "PKGEMAIL110.DATAID_1101"
        Dim cmd65 As New OracleCommand(sql, conn)
        cmd65.CommandType = CommandType.StoredProcedure
        cmd65.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd65.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1114")      'CGTMSE Accounts Not Linked With CGPAN

        'sql = "PKGEMAIL111.DATAID_1114"
        'Dim cmd66 As New OracleCommand(sql, conn)
        'cmd66.CommandType = CommandType.StoredProcedure
        'cmd66.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd66.ExecuteNonQuery()

        processmessage("Package - Data ID - 1131")      'Clientele Base

        sql = "PKGEMAIL113.DATAID_1131"
        Dim cmd67 As New OracleCommand(sql, conn)
        cmd67.CommandType = CommandType.StoredProcedure
        cmd67.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd67.ExecuteNonQuery()

        processmessage("Package - Data ID - 1141")      'Signature Scanning Pending Report

        sql = "PKGEMAIL114.DATAID_1141"
        Dim cmd68 As New OracleCommand(sql, conn)
        cmd68.CommandType = CommandType.StoredProcedure
        cmd68.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd68.ExecuteNonQuery()

        processmessage("Package - Data ID - 1142")      'Gold Overdue Report

        sql = "PKGEMAIL114.DATAID_1142"
        Dim cmd69 As New OracleCommand(sql, conn)
        cmd69.CommandType = CommandType.StoredProcedure
        cmd69.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd69.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1143")      'Gold NPA In Out 

        'sql = "PKGEMAIL114.DATAID_1143"
        'Dim cmd72 As New OracleCommand(sql, conn)
        'cmd72.CommandType = CommandType.StoredProcedure
        'cmd72.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd72.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1144")      'Tech Product Campaign

        'sql = "PKGEMAIL114.DATAID_1144"
        'Dim cmd71 As New OracleCommand(sql, conn)
        'cmd71.CommandType = CommandType.StoredProcedure
        'cmd71.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd71.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1151")      'SASL Pooling Account

        'sql = "PKGEMAIL115.DATAID_1151"
        'Dim cmd73 As New OracleCommand(sql, conn)
        'cmd73.CommandType = CommandType.StoredProcedure
        'cmd73.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd73.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1152")      'Gold Overdue Report - 2nd

        'sql = "PKGEMAIL115.DATAID_1152"
        'Dim cmd74 As New OracleCommand(sql, conn)
        'cmd74.CommandType = CommandType.StoredProcedure
        'cmd74.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd74.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1153")      'Grihodaya Campaign

        'sql = "PKGEMAIL115.DATAID_1153"
        'Dim cmd75 As New OracleCommand(sql, conn)
        'cmd75.CommandType = CommandType.StoredProcedure
        'cmd75.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd75.ExecuteNonQuery()

        processmessage("Package - Data ID - 1154")      'DBTL Registration Status

        sql = "PKGEMAIL115.DATAID_1154"
        Dim cmd77 As New OracleCommand(sql, conn)
        cmd77.CommandType = CommandType.StoredProcedure
        cmd77.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd77.ExecuteNonQuery()

        If RptDate < "01-04-2016" Then

            processmessage("Package - Data ID - 1191")      'KISAN SAMRIDHI CAMPAIGN

            sql = "PKGEMAIL119.DATAID_1191"
            Dim cmd76 As New OracleCommand(sql, conn)
            cmd76.CommandType = CommandType.StoredProcedure
            cmd76.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd76.ExecuteNonQuery()

        End If

        If RptDate < "30-09-2015" Then

            processmessage("Package - Data ID - 1192")      'JAN SURAKSHA CAMPAIGN

            sql = "PKGEMAIL119.DATAID_1192"
            Dim cmd78 As New OracleCommand(sql, conn)
            cmd78.CommandType = CommandType.StoredProcedure
            cmd78.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd78.ExecuteNonQuery()

        End If

        'processmessage("Package - Data ID - 1123")      'PMJDY Account Status

        'sql = "PKGEMAIL123.DATAID_1231"
        'Dim cmd79 As New OracleCommand(sql, conn)
        'cmd79.CommandType = CommandType.StoredProcedure
        'cmd79.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd79.ExecuteNonQuery()

        If RptDate < "01-02-2016" Then

            processmessage("Package - Data ID - 1185")      'ATM Junior Card Status

            sql = "PKGEMAIL118.DATAID_1185"
            Dim cmd76 As New OracleCommand(sql, conn)
            cmd76.CommandType = CommandType.StoredProcedure
            cmd76.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd76.ExecuteNonQuery()

        End If

        'processmessage("Package - Data ID - 1166")      'BANKERS Account

        'sql = "PKGEMAIL116.DATAID_1166"
        'Dim cmd77 As New OracleCommand(sql, conn)
        'cmd77.CommandType = CommandType.StoredProcedure
        'cmd77.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd77.ExecuteNonQuery()

        'processmessage("Package - Data ID -1051")       'Business Progress Report

        'sql = "PKGEMAIL105.DATAID_1051"
        'Dim cmd5 As New OracleCommand(sql, conn)
        'cmd5.CommandType = CommandType.StoredProcedure
        'cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        'cmd5.ExecuteNonQuery()

        'processmessage("Package - Data ID -1043")       'KGB Day Book

        'sql = "PKGEMAIL104.DATAID_1043"
        'Dim cmd6 As New OracleCommand(sql, conn)
        'cmd6.CommandType = CommandType.StoredProcedure
        'cmd6.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        'cmd6.ExecuteNonQuery()

        'Process.Start("C:\du\SMS_" & RptDate.ToString("ddMMyyyy") & ".txt")

        'Generating EMail
        'Add "Microsoft Outlook 15.0 Object Library" in Project >> Reference >> Com
        'Add the following in declaration part
        'Imports System.Runtime.InteropServices
        'Imports Outlook = Microsoft.Office.Interop.Outlook

        Dim sql1 As String
        Dim sql2 As String
        Dim tmplength As Integer = 0
        Dim substringno As Integer = 0
        Dim substringstart As Integer = 1
        Dim Bodydata As String = ""

        tempcount = 0
        sql = "SELECT MAIL_DATAID,MAIL_DATASUBID,MAIL_TO,NVL(MAIL_CC,'A')MAIL_CC,NVL(MAIL_BCC,'A')MAIL_BCC,MAIL_SUBJECT FROM Z_EMAIL ORDER BY MAIL_DATAID, MAIL_DATASUBID"
        Dim cmd21 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd21.ExecuteReader()
        While dr.Read()
            Bodydata = ""
            tmplength = 0
            substringno = 0
            substringstart = 1

            sql1 = "SELECT LENGTH(MAIL_BODY) TEMPLENGTH FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
            Dim cmd22 As New OracleCommand(sql1, conn)
            Dim dr1 As OracleDataReader = cmd22.ExecuteReader()
            While dr1.Read()
                tmplength = dr1.Item("TEMPLENGTH")
                Do Until tmplength = 0

                    If tmplength >= 3000 Then

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & ",3000) ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()

                    Else

                        substringno = tmplength

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & "," & substringno & ") ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()
                        tmplength = 0

                    End If

                Loop

            End While
            dr1.Close()

            Dim oApp As Outlook._Application
            oApp = New Outlook.Application()
            Dim outlooksendfromaccount As String
            Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

            If Val(dr.Item("MAIL_DATAID")) >= 1012 And Val(dr.Item("MAIL_DATAID")) <= 1074 And Val(dr.Item("MAIL_DATAID")) <> 1014 Then
                outlooksendfromaccount = "mis@kgbmis.in"
            ElseIf Val(dr.Item("MAIL_DATAID")) >= 1076 Then
                outlooksendfromaccount = "mis@kgbmis.in"
            End If

            'If Val(dr.Item("MAIL_DATAID")) = 1051 Then
            '    outlooksendfromaccount = "mis@kgbmis.in"
            'Else
            '    outlooksendfromaccount = "mis@kgbmis.in"
            'End If

            Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)
            newMail.To = dr.Item("MAIL_TO")
            If dr.Item("MAIL_CC") <> "A" Then
                newMail.CC = dr.Item("MAIL_CC")
            End If
            If dr.Item("MAIL_BCC") <> "A" Then
                newMail.BCC = dr.Item("MAIL_BCC")
            End If
            newMail.Subject = dr.Item("MAIL_SUBJECT")
            newMail.HTMLBody = Bodydata
            newMail.SendUsingAccount = account
            newMail.Send()
            tempcount = tempcount + 1

            processmessage("Sending Mail No - " & tempcount)

        End While
        dr.Close()
        processmessage("")

        MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Sub option61()       'KYC Upload Statistics

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()


        processmessage("Package - KYC Upload Statistcs.")      'KYC

        sql = "PKGBATCHKYCUPDATE.KYCEMAIL"
        Dim cmd8 As New OracleCommand(sql, conn)
        cmd8.CommandType = CommandType.StoredProcedure
        cmd8.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd8.ExecuteNonQuery()

        Dim sql1 As String
        Dim sql2 As String
        Dim tmplength As Integer = 0
        Dim substringno As Integer = 0
        Dim substringstart As Integer = 1
        Dim Bodydata As String = ""

        tempcount = 0
        sql = "SELECT MAIL_DATAID,MAIL_DATASUBID,MAIL_TO,NVL(MAIL_CC,'A')MAIL_CC,NVL(MAIL_BCC,'A')MAIL_BCC,MAIL_SUBJECT FROM Z_EMAIL ORDER BY MAIL_DATAID, MAIL_DATASUBID"
        Dim cmd21 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd21.ExecuteReader()
        While dr.Read()
            Bodydata = ""
            tmplength = 0
            substringno = 0
            substringstart = 1

            sql1 = "SELECT LENGTH(MAIL_BODY) TEMPLENGTH FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
            Dim cmd22 As New OracleCommand(sql1, conn)
            Dim dr1 As OracleDataReader = cmd22.ExecuteReader()
            While dr1.Read()
                tmplength = dr1.Item("TEMPLENGTH")
                Do Until tmplength = 0

                    If tmplength >= 3000 Then

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & ",3000) ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()

                    Else

                        substringno = tmplength

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & "," & substringno & ") ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()
                        tmplength = 0

                    End If

                Loop

            End While
            dr1.Close()

            Dim oApp As Outlook._Application
            oApp = New Outlook.Application()
            Dim outlooksendfromaccount As String
            Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

            outlooksendfromaccount = "mis@kgbmis.in"
            'If Val(dr.Item("MAIL_DATAID")) >= 1012 And Val(dr.Item("MAIL_DATAID")) <= 1074 And Val(dr.Item("MAIL_DATAID")) <> 1014 Then
            '    outlooksendfromaccount = "mis@kgbmis.in"
            'ElseIf Val(dr.Item("MAIL_DATAID")) >= 1076 Then
            '    outlooksendfromaccount = "mis@kgbmis.in"
            '    'outlooksendfromaccount = "sudhi.kms@gmail.com"
            'End If

            'If Val(dr.Item("MAIL_DATAID")) = 1051 Then
            '    outlooksendfromaccount = "mis@kgbmis.in"
            'Else
            '    outlooksendfromaccount = "mis@kgbmis.in"
            'End If

            Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)
            newMail.To = dr.Item("MAIL_TO")
            If dr.Item("MAIL_CC") <> "A" Then
                newMail.CC = dr.Item("MAIL_CC")
            End If
            If dr.Item("MAIL_BCC") <> "A" Then
                newMail.BCC = dr.Item("MAIL_BCC")
            End If
            newMail.Subject = dr.Item("MAIL_SUBJECT")
            newMail.HTMLBody = Bodydata
            newMail.SendUsingAccount = account
            newMail.Send()
            tempcount = tempcount + 1

            processmessage("Sending Mail No - " & tempcount)

        End While
        dr.Close()
        processmessage("")

        processmessage("Package - KYC SMS Creation.")      'SMS

        sql = "PKGBATCHKYCUPDATE.SMSCREATION"
        Dim cmd18 As New OracleCommand(sql, conn)
        cmd18.CommandType = CommandType.StoredProcedure
        cmd18.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd18.ExecuteNonQuery()

        processmessage("Creating Weekly.txt")

        tempvar = ""
        Dim sw1 As StreamWriter = New StreamWriter("c:/du/Weekly.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 2 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd12 As New OracleCommand(sql, conn)
        Dim dr4 As OracleDataReader = cmd12.ExecuteReader()
        While dr4.Read()
            tempvar = dr4.Item("REPORTDATA")
            sw1.WriteLine(tempvar)
        End While
        dr4.Close()
        sw1.Close()

        processmessage("Creating Cumulative.txt")

        tempvar = ""
        Dim sw2 As StreamWriter = New StreamWriter("c:/du/Cumulative.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 3 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd13 As New OracleCommand(sql, conn)
        Dim dr5 As OracleDataReader = cmd13.ExecuteReader()
        While dr5.Read()
            tempvar = dr5.Item("REPORTDATA")
            sw2.WriteLine(tempvar)
        End While
        dr5.Close()
        sw2.Close()

        MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Sub option43()       'Figures At A Glance

        ' Checking whether email.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\email.txt"
        'file2 = "c:\du\40102.misdo"

        checkfile(file1, "Rename the EMail file 40101.email as email.txt and place in c:/du folder")
        'checkfile(file2, "Place Misdo file 40102.misdo in c:/du folder")

        processmessage("Validating date")
        tempvar1 = readNthLine(file1, 0)

        Try

            tempdate1 = CDate(tempvar1)

        Catch ex As Exception

            MsgBox("Invalid date in first line of email.txt file", MsgBoxStyle.Critical, "Invalid date")
            Exit Sub

        End Try

        If RptDate <> tempdate1 Then

            MsgBox("Report date and date in email.txt do not match", MsgBoxStyle.Critical, "Mismatch in date")
            Exit Sub

        End If

        uploadfiledata(file1, username, "Y")
        'uploadfiledata(file2, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - Get Data")

        sql = "PKGEMAIL102.GETDATA"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Package - Data ID - 1181")      'Figures At Glance

        sql = "PKGEMAIL118.DATAID_1181"
        Dim cmd9 As New OracleCommand(sql, conn)
        cmd9.CommandType = CommandType.StoredProcedure
        cmd9.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd9.ExecuteNonQuery()

        'processmessage("Package - MISDO_INSERT")        'MISDO Upload

        'sql = "PKGMISDOUPLOAD.MISDO_INSERT"
        'Dim cmd5 As New OracleCommand(sql, conn)
        'cmd5.CommandType = CommandType.StoredProcedure
        'cmd5.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1014")      'Figures At Glance

        'sql = "PKGEMAIL101.DATAID_1014"
        'Dim cmd9 As New OracleCommand(sql, conn)
        'cmd9.CommandType = CommandType.StoredProcedure
        'cmd9.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd9.ExecuteNonQuery()

        'Generating EMail
        'Add "Microsoft Outlook 15.0 Object Library" in Project >> Reference >> Com
        'Add the following in declaration part
        'Imports System.Runtime.InteropServices
        'Imports Outlook = Microsoft.Office.Interop.Outlook

        Dim sql1 As String
        Dim sql2 As String
        Dim tmplength As Integer = 0
        Dim substringno As Integer = 0
        Dim substringstart As Integer = 1
        Dim Bodydata As String = ""

        tempcount = 0
        sql = "SELECT MAIL_DATAID,MAIL_DATASUBID,MAIL_TO,NVL(MAIL_CC,'A')MAIL_CC,NVL(MAIL_BCC,'A')MAIL_BCC,MAIL_SUBJECT FROM Z_EMAIL ORDER BY MAIL_DATAID, MAIL_DATASUBID"
        Dim cmd21 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd21.ExecuteReader()
        While dr.Read()
            Bodydata = ""
            tmplength = 0
            substringno = 0
            substringstart = 1

            sql1 = "SELECT LENGTH(MAIL_BODY) TEMPLENGTH FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
            Dim cmd22 As New OracleCommand(sql1, conn)
            Dim dr1 As OracleDataReader = cmd22.ExecuteReader()
            While dr1.Read()
                tmplength = dr1.Item("TEMPLENGTH")
                Do Until tmplength = 0

                    If tmplength >= 3000 Then

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & ",3000) ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()

                    Else

                        substringno = tmplength

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & "," & substringno & ") ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()
                        tmplength = 0

                    End If

                Loop

            End While
            dr1.Close()

            Dim oApp As Outlook._Application
            oApp = New Outlook.Application()
            Dim outlooksendfromaccount As String
            Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

            'If Val(dr.Item("MAIL_DATASUBID")) >= 40401 Then
            '    outlooksendfromaccount = "smgbmis@gmail.com"
            'Else
            '    outlooksendfromaccount = "smgbmis1@gmail.com"
            'End If

            outlooksendfromaccount = "fag@kgbmis.in"

            Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)
            newMail.To = dr.Item("MAIL_TO")
            If dr.Item("MAIL_CC") <> "A" Then
                newMail.CC = dr.Item("MAIL_CC")
            End If
            If dr.Item("MAIL_BCC") <> "A" Then
                newMail.BCC = dr.Item("MAIL_BCC")
            End If
            newMail.Subject = dr.Item("MAIL_SUBJECT")
            newMail.HTMLBody = Bodydata
            newMail.SendUsingAccount = account
            newMail.Send()
            tempcount = tempcount + 1

            processmessage("Sending Mail No - " & tempcount)

        End While
        dr.Close()
        processmessage("")

        MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Sub option3()       'Upload Files

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

            processmessage("")
            MsgBox("Data of " & totalfiles & " files uploaded successfully", MsgBoxStyle.Information, "Process Completed")

        End If

    End Sub

    Sub option4()       'Upload Files

        ' Checking whether tabdata.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\tabdata.txt"

        checkfile(file1, "Parameter file not found in c:/du folder named as 'tabdata.txt'")

        uploadfiledata(file1, username, "Y")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table c_tempdata")
        oracle_execute_non_query("ten", username, username, "truncate table c_misprint")
        oracle_execute_non_query("ten", username, username, "truncate table c_misadv")
        oracle_execute_non_query("ten", username, username, "truncate table c_misdep")
        oracle_execute_non_query("ten", username, username, "INSERT INTO C_TEMPDATA (TD_PROCESSID,TD_USERID,TD_NUMBER1,TD_MEMO1) SELECT 'TABDATA','FRAN1875',LINENO,LINEDATA FROM Z_DU")

        processmessage("Executing package")

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "pkgdataimport.tabdata"
        Dim cmd46 As New OracleCommand(sql, conn)
        cmd46.CommandType = CommandType.StoredProcedure
        cmd46.Parameters.Add("USERID", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "FRAN1875"
        cmd46.Parameters.Add("TRIALFINALFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "F"
        cmd46.ExecuteNonQuery()

        processmessage("Creating output file")

        ' Creating new file

        Dim file3 As String = "c:\du\tabdata_output.txt"

        If File.Exists(file3) Then

            File.Delete(file3)

        End If

        Dim sw As StreamWriter = New StreamWriter(file3)
        sql = "select reportdata from c_misprint order by serialno,subserialno"
        Dim cmd47 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd47.ExecuteReader()
        While dr.Read()
            tempvar = dr.Item("reportdata")
            sw.WriteLine(tempvar)
        End While
        dr.Close()
        sw.Close()

        Dim p As New System.Diagnostics.Process
        Dim s As New System.Diagnostics.ProcessStartInfo("c:\du\tabdata_output.txt")
        s.UseShellExecute = True
        s.WindowStyle = ProcessWindowStyle.Normal
        p.StartInfo = s
        p.Start()

        processmessage("")

        ' Closing Oracle Connection

        conn.Close()
        conn.Dispose()

    End Sub

    Sub option5()       'General

        ' Checking whether general.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\general.txt"

        checkfile(file1, "Parameter file not found in c:/du folder named as 'general.txt'")

        uploadfiledata(file1, username, "Y")

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table c_tempdata")
        oracle_execute_non_query("ten", username, username, "truncate table c_misprint")
        oracle_execute_non_query("ten", username, username, "truncate table c_misadv")
        oracle_execute_non_query("ten", username, username, "truncate table c_misdep")
        oracle_execute_non_query("ten", username, username, "INSERT INTO C_TEMPDATA (TD_PROCESSID,TD_USERID,TD_NUMBER1,TD_MEMO1) SELECT 'GENERAL','FRAN1875',LINENO,LINEDATA FROM Z_DU")

        processmessage("Executing package")

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "pkgdataimport.general"
        Dim cmd46 As New OracleCommand(sql, conn)
        cmd46.CommandType = CommandType.StoredProcedure
        cmd46.Parameters.Add("USERID", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "FRAN1875"
        cmd46.Parameters.Add("TRIALFINALFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "F"
        cmd46.ExecuteNonQuery()

        processmessage("Creating output file")

        ' Creating new file

        Dim file3 As String = "c:\du\general_output.txt"

        If File.Exists(file3) Then

            File.Delete(file3)

        End If

        Dim sw As StreamWriter = New StreamWriter(file3)
        sql = "select reportdata from c_misprint order by serialno,subserialno"
        Dim cmd47 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd47.ExecuteReader()
        While dr.Read()
            tempvar = dr.Item("reportdata")
            sw.WriteLine(tempvar)
        End While
        dr.Close()
        sw.Close()

        Dim p As New System.Diagnostics.Process
        Dim s As New System.Diagnostics.ProcessStartInfo("c:\du\general_output.txt")
        s.UseShellExecute = True
        s.WindowStyle = ProcessWindowStyle.Normal
        p.StartInfo = s
        p.Start()

        processmessage("")

        ' Closing Oracle Connection

        conn.Close()
        conn.Dispose()

    End Sub

    Sub option6()       'Report

        ' Checking whether report.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\report.txt"

        checkfile(file1, "Parameter file not found in c:/du folder named as 'report.txt'")

        uploadfiledata(file1, username, "Y")

        ' Delete existing data

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table c_tempdata")
        oracle_execute_non_query("ten", username, username, "truncate table c_misprint")
        oracle_execute_non_query("ten", username, username, "truncate table c_misadv")
        oracle_execute_non_query("ten", username, username, "truncate table c_misdep")
        oracle_execute_non_query("ten", username, username, "INSERT INTO C_TEMPDATA (TD_PROCESSID,TD_USERID,TD_NUMBER1,TD_MEMO1) SELECT 'REPORT','FRAN1875',LINENO,LINEDATA FROM Z_DU")

        processmessage("Executing package")

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "pkgdataimport.report"
        Dim cmd46 As New OracleCommand(sql, conn)
        cmd46.CommandType = CommandType.StoredProcedure
        cmd46.Parameters.Add("USERID", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "FRAN1875"
        cmd46.Parameters.Add("TRIALFINALFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "F"
        cmd46.ExecuteNonQuery()

        processmessage("Creating output file")

        ' Creating new file

        Dim file3 As String = "c:\du\report_output.txt"

        If File.Exists(file3) Then

            File.Delete(file3)

        End If

        Dim sw As StreamWriter = New StreamWriter(file3)
        sql = "select reportdata from c_misprint order by serialno,subserialno"
        Dim cmd47 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd47.ExecuteReader()
        While dr.Read()
            tempvar = dr.Item("reportdata")
            sw.WriteLine(tempvar)
        End While
        dr.Close()
        sw.Close()

        Dim p As New System.Diagnostics.Process
        Dim s As New System.Diagnostics.ProcessStartInfo("c:\du\report_output.txt")
        s.UseShellExecute = True
        s.WindowStyle = ProcessWindowStyle.Normal
        p.StartInfo = s
        p.Start()

        processmessage("")

        ' Closing Oracle Connection

        conn.Close()
        conn.Dispose()

    End Sub

    Private Sub Button2_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles Button3.Click

        Dim FILE_NAME As String = "C:\HELP\ReadMe.txt"

        If System.IO.File.Exists(FILE_NAME) = True Then

            Process.Start(FILE_NAME)

        Else

            MsgBox("Sorry!!! Help Not Available.")

        End If

    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then

            SendKeys.Send("{tab}")

        End If

        ' To work also set the forms keypreview property to true

    End Sub

    Sub option7()       'KGB Business Progress Report

        ' Checking whether email.txt, nmgb.txt and npa.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\email.txt"
        file2 = "c:\du\nmgb.txt"
        file3 = "c:\du\npa.txt"


        checkfile(file1, "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder")
        checkfile(file2, "Rename the NMGB Trial Balance(TB_XXXXXXXX.xls) as 'nmgb.txt' (Replace tab with |) and place in C:/DU folder")
        checkfile(file3, "Rename the NMGB NPA(NPA_XXXXXXXX.xls) File as 'npa.txt' (Replace tab with |) and place in C:/DU folder")

        processmessage("Validating date")

        tempvar = readNthLine(file1, 0)

        Try

            tempdate = CDate(tempvar)

        Catch ex As Exception

            MsgBox("Invalid date in first line of email.txt file", MsgBoxStyle.Critical, "Invalid date")
            Exit Sub

        End Try

        If RptDate <> tempdate Then

            MsgBox("Report date and date in email.txt do not match", MsgBoxStyle.Critical, "Mismatch in date")
            Exit Sub

        End If

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")
        uploadfiledata(file3, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - Data ID -1051")

        sql = "PKGEMAIL105.DATAID_1051"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd5.ExecuteNonQuery()

        sendemail("mis@kgbmis.in", "ten", username, username)

    End Sub

    Sub option8()       'KGB Day Book

        ' Checking whether SMGBDB.txt and nmgb.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\smgbdb.txt"
        file2 = "c:\du\nmgb.txt"

        checkfile(file1, "Rename the MISDO File (40124_XX-XX-XXXX.misdo) as 'smgbdb.txt' and place in C:/DU folder")
        checkfile(file2, "Rename the NMGB Trial Balance(TB_XXXXXXXX.xls) as 'nmgb.txt' (Replace tab with |) and place in C:/DU folder")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        processmessage("Package - Get Data")

        processmessage("Package - Data ID - 1043")

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGEMAIL104.DATAID_1043"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd5.ExecuteNonQuery()

        sendemail("mis@kgbmis.in", "ten", username, username)

    End Sub

    Sub option9()       'Business Review

        ' Checking whether email2.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\email2.txt"

        checkfile(file1, "Rename the EMail file 40102_XX-XX-XXXX.email as email2.txt and place in c:/du folder")

        uploadfiledata(file1, username, "Y")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        processmessage("Package - Get Data")

        processmessage("Package - Business Review")

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGEMAIL106.DATAID_1061"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, ParameterDirection.Input).Value = "BR"
        cmd5.ExecuteNonQuery()

        'sendemail("smgbmis4@gmail.com", "ten", username, username)
        sendemail("br@kgbmis.in", "ten", username, username)

    End Sub

    Sub option10()      'KGB First - Outstanding

        ' Checking whether MPR_BALANCE_OS.txt, MPR_NO_OS.txt and SMGBFIRST_OS.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\mpr_balance_os.txt"
        file2 = "c:\du\mpr_no_os.txt"
        file3 = "c:\du\smgbfirst_os.txt"

        checkfile(file1, "Create a file in c:/DU named MPR_BALANCE_OS.txt reading NMGB File MPR >> Bal_Amt after converting to pipe delimited format")
        checkfile(file2, "Create a file in c:/DU named MPR_NO_OS.txt reading NMGB File MPR >> Bal_Count after converting to pipe delimited format")
        checkfile(file3, "Place SMGB First Outstanding Bank as a whole (MASRPT 802) into C:\DU\ and rename it as 'SMGBFIRST_OS.txt'")

        uploadfiledata_without_trim(file1, username, "Y")
        uploadfiledata_without_trim(file2, username, "N")
        uploadfiledata_without_trim(file3, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        processmessage("Package - DATAID_1052")

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGEMAIL105.DATAID_1052"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd5.ExecuteNonQuery()

        sendemail("smgbmis2@gmail.com", "ten", username, username)

    End Sub

    Sub option11()      'KGB First - Disbursement

        ' Checking whether MPR_BALANCE_DISB.txt, MPR_NO_DISB.txt and SMGBFIRST_DISB.txt file exists

        processmessage("Checking files")


        file1 = "c:\du\mpr_disb.txt"
        file2 = "c:\du\smgbfirst_disb.txt"

        checkfile(file1, "Create a file in c:/DU named MPR_DISB.txt reading the sheet 'Disb_Upto' in Excel file 'MasterFile.xlsx' after converting to pipe delimited format")
        checkfile(file2, "Place SMGB First Outstanding Bank as a whole (MASRPT 802) into C:\DU\ and rename it as 'SMGBFIRST_DISB.txt'")

        uploadfiledata_without_trim(file1, username, "Y")
        uploadfiledata_without_trim(file2, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        processmessage("Package - DATAID_1053")

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGEMAIL105.DATAID_1053"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd5.ExecuteNonQuery()

        sendemail("smgbmis2@gmail.com", "ten", username, username)

    End Sub

    Sub option12()      'KGB First - NPA

        ' Checking whether MPR_BALANCE_NPA.txt, MPR_NO_NPA.txt and SMGBFIRST_NPA.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\mpr_balance_npa.txt"
        file2 = "c:\du\mpr_no_npa.txt"
        file3 = "c:\du\smgbfirst_npa.txt"

        checkfile(file1, "Create a file in c:/DU named MPR_BALANCE_npa.txt reading NMGB File MPR >> Bal_Amt after converting to pipe delimited format")
        checkfile(file2, "Create a file in c:/DU named MPR_NO_NPA.txt reading NMGB File MPR >> Bal_Count after converting to pipe delimited format")
        checkfile(file3, "Place SMGB First Outstanding Bank as a whole (MASRPT 802) into C:\DU\ and rename it as 'SMGBFIRST_NPA.txt'")

        uploadfiledata_without_trim(file1, username, "Y")
        uploadfiledata_without_trim(file2, username, "N")
        uploadfiledata_without_trim(file3, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        processmessage("Package - DATAID_1054")

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGEMAIL105.DATAID_1054"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd5.ExecuteNonQuery()

        sendemail("smgbmis2@gmail.com", "ten", username, username)

    End Sub

    Sub option13()      'MISDO Upload

        Dim dirs As String() = Directory.GetFiles("c:\du", "*.misdo")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files having extension .misdo exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata(dir, username, "Y")

                Else

                    uploadfiledata(dir, username, "N")

                End If

            Next

            ' Connecting to oracle data base

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - MISDO_INSERT")

            sql = "PKGMISDOUPLOAD.MISDO_INSERT"
            Dim cmd5 As New OracleCommand(sql, conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.ExecuteNonQuery()

            processmessage("")

            MsgBox("Data of " & totalfiles & " files uploaded successfully", MsgBoxStyle.Information, "Process Completed")

            conn.Close()
            conn.Dispose()

        End If

    End Sub

    Sub option14()          'ATM Data Mismatch between Finacle & Switch reports

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - DATAID_1044")

            sql = "PKGEMAIL104.DATAID_1044"
            Dim cmd5 As New OracleCommand(sql, conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.ExecuteNonQuery()

            sendemail("mis@kgbmis.in", "ten", username, username)

        End If

    End Sub

    Sub option15()          'CIBIL Upload File Creation (Live)

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0
        Dim tempvar As String
        Dim Outputfolderpath As String = Disk & ":/CIBIL"

        ' Creating output folder path

        processmessage("Creating output folder path")

        If Directory.Exists(Outputfolderpath) Then

            System.IO.Directory.Delete(Outputfolderpath, True)

        End If

        System.IO.Directory.CreateDirectory(Outputfolderpath)

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata(dir, username, "Y")

                Else

                    uploadfiledata(dir, username, "N")

                End If

            Next

            ' Connecting to oracle data base

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - PKGCIBIL_UPLOAD.CIBIL")

            sql = "PKGCIBIL_UPLOAD.CIBIL"
            Dim cmd5 As New OracleCommand(sql, conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.ExecuteNonQuery()

            processmessage("Creating CIBIL_Non_Individual.txt")

            Dim sl As Integer = 0
            Dim sw As StreamWriter = New StreamWriter(Disk & ":/cibil/BR04640001_" & Today.ToString("ddMMyyyy") & "_NI.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=2 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd11 As New OracleCommand(sql, conn)
            Dim dr As OracleDataReader = cmd11.ExecuteReader()
            While dr.Read()
                sl = sl + 1
                processmessage("Creating CIBIL_Non_Individual.txt Line No:" & sl)
                System.Windows.Forms.Application.DoEvents()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            'processmessage("Creating CIBIL_Individual.txt")

            'tempvar = ""
            'sl = 0
            'sw = New StreamWriter(Disk & ":/cibil/BR04640001_" & Today.ToString("ddMMyyyy") & "_IN.txt")
            'sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=1 ORDER BY SERIALNO,SUBSERIALNO"
            'Dim cmd12 As New OracleCommand(sql, conn)
            'dr = cmd12.ExecuteReader()
            'While dr.Read()
            '    sl = sl + 1
            '    processmessage("Creating CIBIL_Individual.txt Line No:" & sl)
            '    System.Windows.Forms.Application.DoEvents()
            '    tempvar = tempvar & dr.Item("REPORTDATA")
            'End While
            'dr.Close()
            'sw.WriteLine(tempvar)
            'sw.Close()

            processmessage("Creating Summary_Non_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/summary_non_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=4 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd15 As New OracleCommand(sql, conn)
            dr = cmd15.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating Summary_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/summary_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=3 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd13 As New OracleCommand(sql, conn)
            dr = cmd13.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating CustUpld_General_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/custupld_general_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=5 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd14 As New OracleCommand(sql, conn)
            dr = cmd14.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating CustUpld_General_Non_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/custupld_general_non_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=6 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd16 As New OracleCommand(sql, conn)
            dr = cmd16.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating Annexure_A_Non_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/annexurea_non_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=7 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd17 As New OracleCommand(sql, conn)
            dr = cmd17.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating Annexure_A_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/annexurea_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=8 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd18 As New OracleCommand(sql, conn)
            dr = cmd18.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()


            processmessage("")

            MsgBox("CIBIL Upload Files Created Successfully", MsgBoxStyle.Information, "Process Completed")

            conn.Close()
            conn.Dispose()

        End If

    End Sub

    Sub option66()          'CIBIL Upload File Creation (Close)

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0
        Dim tempvar As String
        Dim Outputfolderpath As String = Disk & ":/CIBIL"

        ' Creating output folder path

        processmessage("Creating output folder path")

        If Directory.Exists(Outputfolderpath) Then

            System.IO.Directory.Delete(Outputfolderpath, True)

        End If

        System.IO.Directory.CreateDirectory(Outputfolderpath)

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata(dir, username, "Y")

                Else

                    uploadfiledata(dir, username, "N")

                End If

            Next

            ' Connecting to oracle data base

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - PKGCIBIL_UPLOAD.CIBIL")

            sql = "PKGCIBIL_UPLOAD.CIBIL"
            Dim cmd5 As New OracleCommand(sql, conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.ExecuteNonQuery()

            processmessage("Creating CIBIL_Non_Individual.txt")

            Dim sl As Integer = 0
            Dim sw As StreamWriter = New StreamWriter(Disk & ":/cibil/BR04640001_" & Today.ToString("ddMMyyyy") & "_NI.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=2 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd11 As New OracleCommand(sql, conn)
            Dim dr As OracleDataReader = cmd11.ExecuteReader()
            While dr.Read()
                sl = sl + 1
                processmessage("Creating CIBIL_Non_Individual.txt Line No:" & sl)
                System.Windows.Forms.Application.DoEvents()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating CIBIL_Individual.txt")

            tempvar = ""
            sl = 0
            sw = New StreamWriter(Disk & ":/cibil/BR04640001_" & Today.ToString("ddMMyyyy") & "_IN.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=1 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd12 As New OracleCommand(sql, conn)
            dr = cmd12.ExecuteReader()
            While dr.Read()
                sl = sl + 1
                processmessage("Creating CIBIL_Individual.txt Line No:" & sl)
                System.Windows.Forms.Application.DoEvents()
                tempvar = tempvar & dr.Item("REPORTDATA")
            End While
            dr.Close()
            sw.WriteLine(tempvar)
            sw.Close()

            processmessage("Creating Summary_Non_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/summary_non_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=4 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd15 As New OracleCommand(sql, conn)
            dr = cmd15.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating Summary_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/summary_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=3 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd13 As New OracleCommand(sql, conn)
            dr = cmd13.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating CustUpld_General_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/custupld_general_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=5 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd14 As New OracleCommand(sql, conn)
            dr = cmd14.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating CustUpld_General_Non_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/custupld_general_non_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=6 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd16 As New OracleCommand(sql, conn)
            dr = cmd16.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating Annexure_A_Non_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/annexurea_non_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=7 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd17 As New OracleCommand(sql, conn)
            dr = cmd17.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating Annexure_A_Individual.txt")

            sw = New StreamWriter(Disk & ":/cibil/annexurea_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=8 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd18 As New OracleCommand(sql, conn)
            dr = cmd18.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()


            processmessage("")

            MsgBox("CIBIL Upload Files Created Successfully", MsgBoxStyle.Information, "Process Completed")

            conn.Close()
            conn.Dispose()

        End If

    End Sub

    Sub option16()          'EMail Daily Reports

        'Generating EMail
        'Add "Microsoft Outlook 15.0 Object Library" in Project >> Reference >> Com
        'Add the following in declaration part
        'Imports System.Runtime.InteropServices
        'Imports Outlook = Microsoft.Office.Interop.Outlook

        Dim oApp As Outlook._Application
        oApp = New Outlook.Application()
        Dim outlooksendfromaccount As String
        Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        Dim dirs As String() = Directory.GetFiles("c:\temp")
        Dim dir As String

        outlooksendfromaccount = "mis@kgbmis.in"

        Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)

        newMail.To = "smgbmis@gmail.com;FRANKLINKF1.57372850@E2F.SUGARSYNC.COM"
        newMail.Subject = "Daily Reports/Files"
        newMail.HTMLBody = "<html><body><head><style>.normalandleft{TEXT-ALIGN: left; FONT-FAMILY: arial, helvetica, sans-serif; FONT-SIZE: 9pt; FONT-WEIGHT: normal;}</style></head><p class=normalandleft>Dear Sir,</p><p class=normalandleft>Enclosed please find the daily reports/files generated during the day.</p><p class=normalandleft>Yours faithfully,</p><p class=normalandleft>MIS Team</p></body></html>"
        For Each dir In dirs
            newMail.Attachments.Add(dir)
        Next
        newMail.SendUsingAccount = account
        newMail.Send()

        MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Sub option17()      'NPCI Linked Aadhaar - Upload file creation

        ' Checking whether npci_aadhaar.txt file exists

        Dim tempvar As String
        Dim tempcount As String = 0

        processmessage("Checking files")

        file1 = "c:\du\npci_aadhaar.txt"

        checkfile(file1, "Place the report downloaded from NPCI in c:/DU naming as npci_aadhaar.txt")

        uploadfiledata(file1, username, "Y")

        ' Connecting to oracle data base

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        ' Calling packages

        processmessage("Package - PKGMISTOOL2.AADHAR_NPCI_UPLOAD")

        sql = "PKGMISTOOL2.AADHAR_NPCI_UPLOAD"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Creating npci_aadhar_upload.txt")

        tempvar = ""
        Dim lineno As Integer
        Dim fileno As Integer
        lineno = 0
        fileno = 1
        Dim filepath As String
        'output_Wrtr = System.IO.File.CreateText(outputFile)
        filepath = "c:/du/npci_aadhaar_upload1.txt"
        'Dim sw1 As StreamWriter = New StreamWriter(filepath)
        Dim sw1 As StreamWriter
        sw1 = System.IO.File.CreateText(filepath)
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO < 5 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd12 As New OracleCommand(sql, conn)
        Dim dr1 As OracleDataReader = cmd12.ExecuteReader()
        While dr1.Read()
            lineno = lineno + 1
            If lineno = 5000 Then
                lineno = 0
                fileno = fileno + 1
                filepath = "c:/du/npci_aadhaar_upload" & fileno & ".txt"
                sw1 = System.IO.File.CreateText(filepath)
                tempvar = "A"
                sw1.WriteLine(tempvar)
                tempvar = "C_DATACAPT"
                sw1.WriteLine(tempvar)
                tempvar = "DC_DATAID|DC_SOLID|DC_FIELD01|DC_FIELD02|DC_FIELD03|DC_FIELD04|DC_FIELD05|DC_FIELD06|DC_FIELD07|DC_FIELD08|DC_FIELD09|DC_FIELD10|DC_FIELD11|DC_FIELD12|DC_FIELD13|DC_FIELD14|DC_FIELD15|DC_FIELD16|DC_FIELD17|DC_FIELD18|DC_FIELD19|DC_FIELD20|DC_CUSERID|DC_CDATE|DC_MUSERID|DC_MDATE|DC_DUSERID|DC_DDATE|DC_VUSERID|DC_VDATE|DC_DATE01|DC_DATE02|DC_DATE03|DC_DATE04|DC_DATE05|DC_DATE06|DC_DATE07|DC_DATE08|DC_DATE09|DC_DATE10|DC_NUMBER01|DC_NUMBER02|DC_NUMBER03|DC_NUMBER04|DC_NUMBER05|DC_NUMBER06|DC_NUMBER07|DC_NUMBER08|DC_NUMBER09|DC_NUMBER10"
                sw1.WriteLine(tempvar)
            End If
            tempvar = dr1.Item("REPORTDATA")
            sw1.WriteLine(tempvar)
            sw1.AutoFlush = True
            System.Windows.Forms.Application.DoEvents()
            'System.Windows.Forms.Application.DoEvents()

        End While
        dr1.Close()
        sw1.Close()

        'splitfile2()
        'SplitFile("c:/du/npci_aadhaar_upload.txt", "c:/du/npci_split.txt", 5)

        tempvar = ""
        'Dim sw2 As StreamWriter = New StreamWriter("c:/du/npci_aadhaar_delete.txt")
        lineno = 0
        fileno = 1
        filepath = "c:/du/npci_aadhaar_delete1.txt"
        Dim sw2 As StreamWriter
        sw2 = System.IO.File.CreateText(filepath)

        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO BETWEEN 5 AND 8 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd13 As New OracleCommand(sql, conn)
        Dim dr2 As OracleDataReader = cmd13.ExecuteReader()
        While dr2.Read()
            lineno = lineno + 1
            If lineno = 500 Then
                lineno = 0
                fileno = fileno + 1
                filepath = "c:/du/npci_aadhaar_delete" & fileno & ".txt"
                sw2 = System.IO.File.CreateText(filepath)
                tempvar = "U"
                sw2.WriteLine(tempvar)
                tempvar = "C_DATACAPT"
                sw2.WriteLine(tempvar)
                tempvar = "DC_DATAID|DC_NUMBER02$DC_DATE02|DC_DATE01"
                sw2.WriteLine(tempvar)
            End If
            tempvar = dr2.Item("REPORTDATA")
            sw2.WriteLine(tempvar)
            sw2.AutoFlush = True
            System.Windows.Forms.Application.DoEvents()

        End While
        dr2.Close()
        sw2.Close()

        tempvar = ""

        Dim sw3 As StreamWriter = New StreamWriter("c:/du/npci_aadhaar_Status.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO > 8 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd14 As New OracleCommand(sql, conn)
        Dim dr3 As OracleDataReader = cmd14.ExecuteReader()
        While dr3.Read()
            tempvar = dr3.Item("REPORTDATA")

            sw3.WriteLine(tempvar)
        End While
        dr3.Close()
        sw3.Close()

        processmessage("")

        MsgBox("Upload file created successfully", MsgBoxStyle.Information, "Process Completed")

        conn.Close()
        conn.Dispose()

    End Sub
    Sub splitfile2()
        Dim sb As New Text.StringBuilder
        Dim directory As String = "c:/du/npci_aadhaar_upload.txt"
        Dim sr As StreamReader = New StreamReader(directory)
        Dim tempvarArray() As String
        Dim tempvar As String
        Dim tempvar1 As String
        Dim directory1 As String
        Dim lineno As Integer
        Dim fileno As Integer
        fileno = 1
        lineno = 0
        tempvar1 = ""
        Do While sr.Peek() >= 0
            directory1 = "c:/du/npci" & fileno & ".txt"
            tempvar = sr.ReadLine()
            tempvar1 = tempvar1 & tempvar
            tempvarArray = Split(tempvar1, "abcd")
            If lineno > 5000 Then
                lineno = 0
                fileno = fileno + 1
                File.WriteAllLines(directory1, tempvarArray)
            End If
            lineno = lineno + 1
            'File.WriteAllLines(directory1, tempvarArray)
        Loop
    End Sub
    Private Function SplitFile( _
    ByVal inputFileName As String, ByVal outputFileName As String, ByVal numberOfFiles As Integer) _
    As List(Of String)
        Dim returnList As New List(Of String)
        Try
            Dim outputFileExtension As String = IO.Path.GetExtension(outputFileName)
            outputFileName = outputFileName.Replace(outputFileExtension, "")
            Dim sr As New IO.StreamReader(inputFileName)
            Dim fileLength As Long = sr.BaseStream.Length
            Dim baseBufferSize As Integer = CInt(fileLength \ numberOfFiles)
            Dim finished As Boolean = False
            Dim fileCount As Integer = 1
            Do Until finished
                Dim bufferSize As Integer = baseBufferSize
                Dim originalPosition As Long = sr.BaseStream.Position
                'find line first line feed after the base buffer length
                sr.BaseStream.Position += bufferSize
                If sr.BaseStream.Position < fileLength Then
                    Do Until sr.Read = 10
                        bufferSize += 1
                    Loop
                    bufferSize += 1
                Else
                    bufferSize = CInt(fileLength - originalPosition)
                    finished = True
                End If
                'write the chunk of data to a buffer in memory
                sr.BaseStream.Position = originalPosition
                Dim buffer(bufferSize - 1) As Byte
                sr.BaseStream.Read(buffer, 0, bufferSize)
                'write the chunk of data to a file
                Dim outputPath As String = outputFileName & fileCount.ToString & outputFileExtension
                returnList.Add(outputPath)
                My.Computer.FileSystem.WriteAllBytes( _
                outputPath, buffer, False)
                fileCount += 1
            Loop
        Catch ex As Exception
            Console.Write(ex.ToString)
        End Try
        Return returnList
    End Function


    Sub option18()      'Day end eMails

        ' Checking whether 40998,40995,40994,KYC.TXT files exists

        processmessage("Checking files")

        file1 = "c:\du\40994.txt"
        file2 = "c:\du\40995.txt"
        file3 = "c:\du\40998.txt"
        file4 = "c:\du\KYC.txt"
        file5 = "c:\du\40991.txt"

        checkfile(file1, "Rename the file 40994_XX-XX-XXXX_AC1.TXT as 40994.TXT and place in c:/du folder")
        checkfile(file2, "Rename the file 40995_XX-XX-XXXX_AC1.TXT as 40995.TXT and place in c:/du folder")
        checkfile(file3, "Rename the file 40998AC1.TXT as 40998.TXT and place in c:/du folder")
        checkfile(file4, "Rename the upload error file KYC_XXXXXX.TXT as KYC.TXT and place in c:/du folder")
        checkfile(file5, "Rename the file 40991_XX-XX-XXXX_AC1.TXT as 40991.TXT and place in c:/du folder")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")
        uploadfiledata(file3, username, "N")
        uploadfiledata(file4, username, "N")
        uploadfiledata(file5, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - PKGEMAIL107.DATAID_1071")

        sql = "PKGEMAIL107.DATAID_1071"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Package - PKGEMAIL107.DATAIID_1072")

        sql = "PKGEMAIL107.DATAID_1072"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.ExecuteNonQuery()

        processmessage("Package - PKGEMAIL107.DATAIID_1077")

        sql = "PKGEMAIL107.DATAID_1077"
        Dim cmd6 As New OracleCommand(sql, conn)
        cmd6.CommandType = CommandType.StoredProcedure
        cmd6.ExecuteNonQuery()

        processmessage("Package - PKGEMAIL107.DATAIID_1078")

        sql = "PKGEMAIL107.DATAID_1078"
        Dim cmd7 As New OracleCommand(sql, conn)
        cmd7.CommandType = CommandType.StoredProcedure
        cmd7.ExecuteNonQuery()


        'sendemail("smgbmis3@gmail.com", "ten", username, username)
        sendemail("mis@kgbmis.in", "ten", username, username)

    End Sub

    Sub option48()      'Branch Intimation Letter

        ' Checking whether files exists in C:/DU folder

        processmessage("Checking files")

        file1 = "c:\du\DATA.txt"

        checkfile(file1, "Rename the data file as DATA.TXT and place in c:/du folder")

        uploadfiledata(file1, username, "Y")

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - PKGEMAIL113.DATAID_1136")

        sql = "PKGEMAIL113.DATAID_1136"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        sendemail("dipsdot@gmail.com", "ten", username, username)

    End Sub

    Sub option53()      'PMJJBY/PMSBY/APY Enrollment Status

        processmessage("Checking files")

        file1 = "c:\du\Data.txt"
        file2 = "c:\du\Data1.txt"

        checkfile(file1, "Rename the data file as Data.TXT and place in c:/du folder")
        checkfile(file2, "Rename the data file as Data1.TXT and place in c:/du folder")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - PKGEMAIL118.DATAID_1182")

        sql = "PKGEMAIL118.DATAID_1182"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd4.ExecuteNonQuery()

        processmessage("Package - PKGEMAIL118.DATAID_1186")

        sql = "PKGEMAIL118.DATAID_1186"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd5.ExecuteNonQuery()

        'processmessage("Package - PKGEMAIL116.DATAID_1161")

        'sql = "PKGEMAIL116.DATAID_1161"
        'Dim cmd4 As New OracleCommand(sql, conn)
        'cmd4.CommandType = CommandType.StoredProcedure
        'cmd4.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd4.ExecuteNonQuery()

        'sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'SARFAESI' ORDER BY SOLID"
        'display_in_File(sql, "C:\du\SMS_SARFAESI.txt")
        'Process.Start("C:\du\SMS_SARFAESI.txt")

        'processmessage("Package - PKGEMAIL116.DATAID_1162")

        'sql = "PKGEMAIL116.DATAID_1162"
        'Dim cmd5 As New OracleCommand(sql, conn)
        'cmd5.CommandType = CommandType.StoredProcedure
        'cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        'cmd5.ExecuteNonQuery()

        'sendemail("kgbmis1@gmail.com", "ten", username, username)
        'sendemail("dipsdot@gmail.com", "ten", username, username)
        sendemail("br@kgbmis.in", "ten", username, username)


    End Sub

    Sub option60()      'PROGRESS REPORT AS PER CIRCULAR: 74/2015

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

        End If

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - PKGEMAIL118.DATAID_1183")

        sql = "PKGEMAIL118.DATAID_1183"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd4.ExecuteNonQuery()

        processmessage("Package - PKGEMAIL118.DATAID_1184")

        sql = "PKGEMAIL118.DATAID_1184"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd5.ExecuteNonQuery()

        'sendemail("kgbmis1@gmail.com", "ten", username, username)
        'sendemail("dipsdot@gmail.com", "ten", username, username)
        sendemail("br@kgbmis.in", "ten", username, username)

    End Sub

    Sub option19()          'Business Review - Files to RO

        processmessage("Checking files")

        file1 = Disk & ":\Business Review Report.rar"

        checkfile(file1, "Compressed file Business Review Report.rar not found in D Drive")

        'Generating EMail
        'Add "Microsoft Outlook 15.0 Object Library" in Project >> Reference >> Com
        'Add the following in declaration part
        'Imports System.Runtime.InteropServices
        'Imports Outlook = Microsoft.Office.Interop.Outlook

        Dim oApp As Outlook._Application
        oApp = New Outlook.Application()
        Dim outlooksendfromaccount As String
        Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        Dim dirs As String() = Directory.GetFiles("c:\temp")

        outlooksendfromaccount = "br@kgbmis.in"

        Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)

        newMail.To = "chairmankeralagb@gmail.com;nkkrishnankutty46876@gmail.com;srnair32474@gmail.com;haridasanv@gmail.com"
        newMail.CC = "nmgbrotly@gmail.com;roekm.kgb@gmail.com;nmgbksd@gmail.com;nmgbknrao@gmail.com;rotvm.kgb@gmail.com;roekm.kgb@gmail.com;roktm.kgb@gmail.com;rotsr.kgb@gmail.com;ropma.kgb@gmail.com;rokzd.kgb@gmail.com;roknr.kgb@gmail.com;rokpt.kgb@gmail.com;roksd.kgb@gmail.com;rotly.kgb@gmail.com;PDWing.KGB@gmail.com;kgbitw@gmail.com"
        newMail.BCC = "kgbhomis@gmail.com;franklinkf@gmail.com;udayakumarcv@gmail.com;sureshsmgb@gmail.com;"
        newMail.Subject = "BUSINESS REVIEW - EXCEL AND MAIL MERGE WORD FILE WITH DATA AS ON " & txtdate.Text
        newMail.HTMLBody = "<html><body><head><style>.normalandleft{TEXT-ALIGN: left; FONT-FAMILY: arial, helvetica, sans-serif; FONT-SIZE: 9pt; FONT-WEIGHT: normal;}</style></head><p class=normalandleft>Dear Sir,</p><p class=normalandleft>Enclosed please find the following files containing the business figures of KGB branches as on " & txtdate.Text & ":</p><p class=normalandleft>1. Business Review.xlsx - To view the figures by providing the branch code/RO Code<br>2. Business Review.docx - To print the figures of branches in batch using the inbuilt mail merge facility.<br>3. Business Review Data.txt - Data source for the mail merge word file.  No specific use with that file<br></p><p class=normalandleft>To view/print the data, Download the attachment (compressed file), extract it and place the files in " & Disk & ":\Business Review Report</p><p class=normalandleft>In addition to this, the following facilites are available:</p><p class=normalandleft>1. Business review figures of every Fridays are emailed to branches/RO/HO on the next day<br>2. Business review figure of any day is available in Finacle MIS Server under Report ID - HMISRPT 210<br><p class=normalandleft>Yours faithfully,</p><p class=normalandleft>MIS Team</p></body></html>"
        newMail.Attachments.Add(file1)
        newMail.SendUsingAccount = account
        newMail.Send()

        MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Sub option20()      'KGB Aadhar Enrolled Status

        ' Checking whether EMAIL.txt and NMGB_AADHAR.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\email.txt"
        file2 = "c:\du\nmgb_aadhar.txt"

        checkfile(file1, "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder")
        checkfile(file2, "Create a file in c:/DU named NMGB_AADHAR.txt reading NMGB File AADHARMAPPED.xls after converting to pipe delimited format")

        processmessage("Validating date")

        tempvar = readNthLine(file1, 0)

        Try

            tempdate = CDate(tempvar)

        Catch ex As Exception

            MsgBox("Invalid date in first line of email.txt file", MsgBoxStyle.Critical, "Invalid date")
            Exit Sub

        End Try

        If RptDate <> tempdate Then

            MsgBox("Report date and date in email.txt do not match", MsgBoxStyle.Critical, "Mismatch in date")
            Exit Sub

        End If

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - Get Data")

        sql = "PKGEMAIL102.GETDATA"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Package - Data ID - 1075")

        sql = "PKGEMAIL107.DATAID_1075"
        Dim cmd52 As New OracleCommand(sql, conn)
        cmd52.CommandType = CommandType.StoredProcedure
        cmd52.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd52.ExecuteNonQuery()

        sendemail("mis@kgbmis.in", "ten", username, username)

    End Sub

    Sub option21()          'KGB Daily Reports

        ' Checking whether email.txt, nmgb.txt , npa.txt , smgbd.txt , nmgb_aadhar.txt files exists

        processmessage("Checking files")

        file1 = "c:\du\email.txt"
        file2 = "c:\du\nmgb.txt"
        file3 = "c:\du\npa.txt"
        file4 = "c:\du\smgbdb.txt"
        file5 = "c:\du\nmgb_aadhar.txt"
        file6 = "c:\du\friday.txt"


        checkfile(file1, "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder")
        checkfile(file2, "Rename the NMGB Trial Balance(TB_XXXXXXXX.xls) as 'nmgb.txt' (Replace tab with |) and place in C:/DU folder")
        checkfile(file3, "Rename the NMGB NPA(NPA_XXXXXXXX.xls) File as 'npa.txt' (Replace tab with |) and place in C:/DU folder")
        checkfile(file4, "Rename the MISDO File (40124_XX-XX-XXXX.misdo) as 'smgbdb.txt' and place in C:/DU folder")
        checkfile(file5, "Create a file in c:/DU named NMGB_AADHAR.txt reading NMGB File AADHARMAPPED.xls after converting to pipe delimited format")
        checkfile(file6, "Place last friday file in C:/DU folder as 'FRIDAY.txt'")

        processmessage("Validating date")

        tempvar = readNthLine(file1, 0)

        Try

            tempdate = CDate(tempvar)

        Catch ex As Exception

            MsgBox("Invalid date in first line of email.txt file", MsgBoxStyle.Critical, "Invalid date")
            Exit Sub

        End Try

        If RptDate <> tempdate Then

            MsgBox("Report date and date in email.txt do not match", MsgBoxStyle.Critical, "Mismatch in date")
            Exit Sub

        End If

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")
        uploadfiledata(file3, username, "N")
        uploadfiledata(file4, username, "N")
        uploadfiledata(file5, username, "N")
        uploadfiledata(file6, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - Data ID - 1051")

        sql = "PKGEMAIL105.DATAID_1051"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd5.ExecuteNonQuery()

        processmessage("Package - Data ID - 1043")

        sql = "PKGEMAIL104.DATAID_1043"
        Dim cmd6 As New OracleCommand(sql, conn)
        cmd6.CommandType = CommandType.StoredProcedure
        cmd6.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd6.ExecuteNonQuery()

        processmessage("Package - Get Data")

        sql = "PKGEMAIL102.GETDATA"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Package - Data ID - 1075")

        sql = "PKGEMAIL107.DATAID_1075"
        Dim cmd52 As New OracleCommand(sql, conn)
        cmd52.CommandType = CommandType.StoredProcedure
        cmd52.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd52.ExecuteNonQuery()

        sendemail("smgbmis2@gmail.com", "ten", username, username)

    End Sub

    'Sub option21()          'KGB Daily Reports

    '    ' Checking whether email.txt, nmgb.txt , npa.txt , smgbd.txt , nmgb_aadhar.txt files exists

    '    processmessage("Checking files")

    '    file1 = "c:\du\email.txt"
    '    file2 = "c:\du\friday.txt"


    '    checkfile(file1, "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder")
    '    checkfile(file2, "Place last friday file in C:/DU folder as 'FRIDAY.txt'")

    '    processmessage("Validating date")

    '    tempvar = readNthLine(file1, 0)

    '    Try

    '        tempdate = CDate(tempvar)

    '    Catch ex As Exception

    '        MsgBox("Invalid date in first line of email.txt file", MsgBoxStyle.Critical, "Invalid date")
    '        Exit Sub

    '    End Try

    '    If RptDate <> tempdate Then

    '        MsgBox("Report date and date in email.txt do not match", MsgBoxStyle.Critical, "Mismatch in date")
    '        Exit Sub

    '    End If

    '    uploadfiledata(file1, username, "Y")
    '    uploadfiledata(file2, username, "N")

    '    ' Delete existing data, if any, from c_du table

    '    processmessage("Deleting existing data")

    '    oracle_execute_non_query("ten", username, username, "truncate table z_email")

    '    ' Calling packages

    '    Dim sql As String
    '    Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
    '    Dim conn As New OracleConnection(oradb)
    '    conn.Open()

    '    processmessage("Package - Data ID - 1051")

    '    sql = "PKGEMAIL105.DATAID_1051"
    '    Dim cmd5 As New OracleCommand(sql, conn)
    '    cmd5.CommandType = CommandType.StoredProcedure
    '    cmd5.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
    '    cmd5.ExecuteNonQuery()

    '    processmessage("Package - Get Data")

    '    sql = "PKGEMAIL102.GETDATA"
    '    Dim cmd4 As New OracleCommand(sql, conn)
    '    cmd4.CommandType = CommandType.StoredProcedure
    '    cmd4.ExecuteNonQuery()

    '    sendemail("smgbmis2@gmail.com", "ten", username, username)

    'End Sub

    Sub option9072()      '9072 Insert

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:\du", MsgBoxStyle.Critical, "Error")

        Else


            For Each dir In dirs

                uploadfiledata_without_trim(dir, username, "Y")

                ' Calling packages

                Dim sql As String
                Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
                Dim conn As New OracleConnection(oradb)
                conn.Open()

                processmessage("Inserting data in to tables")

                sql = "PKGMISTOOL2.INSERT_9072"
                Dim cmd52 As New OracleCommand(sql, conn)
                cmd52.CommandType = CommandType.StoredProcedure
                cmd52.ExecuteNonQuery()

                conn.Close()
                conn.Dispose()

            Next

            processmessage("")
            MsgBox("Data of " & totalfiles & " files inserted successfully", MsgBoxStyle.Information, "Process Completed")

        End If

    End Sub
    Sub option9074()          '9074 Insert

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:\du", MsgBoxStyle.Critical, "Error")

        Else

            oracle_execute_non_query("ten", username, username, "truncate table c_misprint")

            For Each dir In dirs

                uploadfiledata_without_trim(dir, username, "Y")

                ' Calling packages

                Dim sql As String
                Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
                Dim conn As New OracleConnection(oradb)
                conn.Open()

                processmessage("Inserting data into tables - " & dir)

                sql = "PKGMISTOOL2.INSERT_9074"
                Dim cmd52 As New OracleCommand(sql, conn)
                cmd52.CommandType = CommandType.StoredProcedure
                cmd52.ExecuteNonQuery()

            Next

            processmessage("")
            MsgBox("Data of " & totalfiles & " files inserted successfully", MsgBoxStyle.Information, "Process Completed")

        End If

    End Sub
    Sub option9071()          '9071 Insert

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:\du", MsgBoxStyle.Critical, "Error")

        Else

            oracle_execute_non_query("ten", username, username, "truncate table c_misprint")

            For Each dir In dirs

                uploadfiledata_without_trim(dir, username, "Y")

                ' Calling packages

                Dim sql As String
                Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
                Dim conn As New OracleConnection(oradb)
                conn.Open()

                processmessage("Inserting data into tables - " & dir)

                sql = "PKGMISTOOL1.INSERT_9071"
                Dim cmd52 As New OracleCommand(sql, conn)
                cmd52.CommandType = CommandType.StoredProcedure
                cmd52.ExecuteNonQuery()

            Next

            processmessage("")
            MsgBox("Data of " & totalfiles & " files inserted successfully", MsgBoxStyle.Information, "Process Completed")

        End If

    End Sub
    'Sub option25()          'Create RO and Branch Folders and convert CIB Files

    '    Dim foldercreationpath As String = "c:\du"
    '    Dim sourcefilepath As String = "C:\DU\CSV"
    '    Dim sourcefileextention As String = "csv"
    '    Dim dirs As String() = Directory.GetFiles(sourcefilepath, "*." & sourcefileextention)
    '    Dim folders As String()
    '    Dim folder As String
    '    Dim dir As String
    '    Dim totalfiles As Integer
    '    Dim tempcount As Integer = 0
    '    Dim filename As String
    '    Dim solid As String
    '    Dim subfolders As String()
    '    Dim subfolder As String
    '    Dim destinationpath As String
    '    'Creating folders and subfolders
    '    createdistrictbranchfolders(foldercreationpath)
    '    totalfiles = dirs.Length
    '    If totalfiles = 0 Then
    '        processmessage("")
    '        MsgBox("No files exists in the folder " & sourcefilepath, MsgBoxStyle.Critical, "Error")
    '    Else
    '        For Each dir In dirs
    '            destinationpath = ""
    '            tempcount = tempcount + 1
    '            filename = GetFileName(dir)
    '            solid = filename.Substring(0, 5)
    '            folders = Directory.GetDirectories(foldercreationpath)
    '            For Each folder In folders
    '                subfolders = Directory.GetDirectories(folder)
    '                For Each subfolder In subfolders
    '                    If InStr(subfolder, solid) > 0 Then
    '                        destinationpath = subfolder
    '                    End If
    '                Next
    '            Next
    '            CreateExcelFromCsvFile(sourcefilepath, filename, sourcefileextention)
    '            processmessage("Converting File No - " & tempcount)
    '            If destinationpath <> "" Then
    '                My.Computer.FileSystem.CopyFile(sourcefilepath & "\" & filename & ".xls", destinationpath & "\" & filename & ".xls", Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
    '                processmessage("Moving File No - " & tempcount)
    '            End If
    '        Next
    '        processmessage("")
    '        MsgBox("Conversion completed successfully", MsgBoxStyle.Information, "Process Completed")
    '    End If
    'End Sub

    Sub option26()          'Create Bank as a whole/All RO's/All Branches report in a single file

        Dim solid As String
        Dim processid As String
        Dim sw As StreamWriter = New StreamWriter("C:\DU\Report.txt")
        processid = InputBox("Enter Process ID :", "Process ID")
        processid = UCase(processid)

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "SELECT 'ALL' RONAME, 'ALL' SOLID FROM DUAL UNION SELECT DISTINCT RONAME, RONAME FROM C_MISONLINEDATE WHERE LENGTH(RONAME) = 5 UNION SELECT RONAME,TO_CHAR(SOLID2) FROM C_MISONLINEDATE WHERE SOLID2 > 40101 ORDER BY 1,2 DESC"
        Dim cmd As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd.ExecuteReader()
        While dr.Read()
            solid = dr("SOLID")

            processmessage("Processing report for SOLID - " & solid)

            sql = "PKGMISTOOL3.PROCESS"
            Dim cmd1 As New OracleCommand(sql, conn)
            cmd1.CommandType = CommandType.StoredProcedure
            cmd1.Parameters.Add("PROCESSID", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = processid
            cmd1.Parameters.Add("SOLID", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = solid
            cmd1.ExecuteNonQuery()

            sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd2 As New OracleCommand(sql, conn)
            Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
            While dr2.Read()
                tempvar = dr2.Item("REPORTDATA").ToString
                If tempvar <> "" Then
                    sw.WriteLine(tempvar)
                End If
            End While
            dr2.Close()

        End While
        sw.Close()
        dr.Close()

        processmessage("")
        MsgBox("Process completed", MsgBoxStyle.Information, "Done")

        conn.Close()
        conn.Dispose()

    End Sub
    Sub option27()          'Get File Names

        Dim foldername As String
        Dim foldernameprintflag As String
        Dim extensionfilter As String
        Dim extension As String
        Dim sw As StreamWriter = New StreamWriter("C:\DU\FileName.txt")
        Dim fileEntries As String()
        Dim fileName As String
        Dim subdirectoryEntries As String()
        Dim subdirectory As String

        foldername = InputBox("Enter folder name :")
        ' foldernameprintflag = InputBox("Print file names only (Y/N):", "", "N")
        foldernameprintflag = InputBox("Required file name structure:" & vbCrLf & vbCrLf & "<Y> With folder name" & vbCrLf & "<N> Without folder name" & vbCrLf & "<S> SQL Execution style", "", "S")

        extensionfilter = InputBox("Enter extension prefixed by dot", "", "ALL")

        fileEntries = Directory.GetFiles(foldername)
        processmessage("Reading folder - " & foldername)
        For Each fileName In fileEntries
            extension = Path.GetExtension(fileName)
            tempvar = fileName
            If UCase(foldernameprintflag) = "N" Then
                tempvar = Path.GetFileName(fileName)
            End If
            If UCase(foldernameprintflag) = "S" Then
                tempvar = "@""" & tempvar & """"

            End If
            If UCase(extensionfilter) = "ALL" Then
                sw.WriteLine(tempvar)
            Else
                If UCase(extension) = UCase(extensionfilter) Then
                    sw.WriteLine(tempvar)
                End If
            End If
        Next fileName

        subdirectoryEntries = Directory.GetDirectories(foldername)
        For Each subdirectory In subdirectoryEntries
            processmessage("Reading folder - " & subdirectory)
            fileEntries = Directory.GetFiles(subdirectory)
            For Each fileName In fileEntries
                extension = Path.GetExtension(fileName)
                tempvar = fileName
                If UCase(foldernameprintflag) = "N" Then
                    tempvar = Path.GetFileName(fileName)
                End If
                If UCase(extensionfilter) = "ALL" Then
                    sw.WriteLine(tempvar)
                Else
                    If UCase(extension) = UCase(extensionfilter) Then
                        sw.WriteLine(tempvar)
                    End If
                End If
            Next fileName
        Next subdirectory

        processmessage("")
        MsgBox("Process completed", MsgBoxStyle.Information, "Done")

        sw.Close()

    End Sub
    Sub uploadfiledata(ByVal filename As String, ByVal username As String, ByVal du_clear_flag As String)

        Dim clob As String
        Dim cloblength As Integer
        Dim slno As Integer

        clob = My.Computer.FileSystem.ReadAllText(filename)
        clob = clob.Replace("'", "`")
        cloblength = clob.Length

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        If du_clear_flag = "Y" Then

            processmessage("Deleting existing data")

            oracle_execute_non_query("ten", username, username, "truncate table z_du")

        End If

        processmessage("Uploading data - " & filename)

        oracle_execute_non_query("ten", username, username, "truncate table z_du_bulk")
        oracle_execute_non_query("ten", username, username, "truncate table z_email")
        tempcount = 0
        slno = 0

        Do Until cloblength = 0

            If cloblength < 3998 Then
                tempvar = clob.Substring(tempcount, cloblength)
                tempcount = tempcount + cloblength
                cloblength = cloblength - cloblength
            Else
                tempvar = clob.Substring(tempcount, 3998)
                tempcount = tempcount + 3998
                cloblength = cloblength - 3998
            End If
            slno = slno + 1
            processmessage1(cloblength)
            sql = "insert into z_du_bulk (slno,bulkdata) values (" & slno & ",'" & tempvar & "')"
            Dim cmd152 As New OracleCommand(sql, conn)
            cmd152.ExecuteNonQuery()

        Loop

        processmessage1("")
        processmessage("Formatting data - " & filename)

        sql = "PKGMISTOOL2.BULKUPLOAD"
        Dim cmd1 As New OracleCommand(sql, conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.Add("FILENAME", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = filename
        cmd1.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()

    End Sub
    Sub uploadfiledata_without_trim(ByVal filename As String, ByVal username As String, ByVal du_clear_flag As String)

        Dim clob As String
        Dim cloblength As Integer
        Dim slno As Integer

        Try

            clob = My.Computer.FileSystem.ReadAllText(filename)
            clob = clob.Replace("'", "`")
            cloblength = clob.Length

        Catch ex As Exception

            clob = ""
            cloblength = 0
            oracle_execute_non_query("ten", username, username, "INSERT INTO C_MISPRINT (SERIALNO,REPORTDATA) VALUES (9999999999,'Visual Studio - Error in reading file " & filename & "')")

        End Try

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        If du_clear_flag = "Y" Then

            processmessage("Deleting existing data")

            oracle_execute_non_query("ten", username, username, "truncate table z_du")

        End If

        processmessage("Uploading data - " & filename)

        oracle_execute_non_query("ten", username, username, "truncate table z_du_bulk")
        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        tempcount = 0
        slno = 0

        Do Until cloblength = 0

            If cloblength < 3998 Then
                tempvar = clob.Substring(tempcount, cloblength)
                tempcount = tempcount + cloblength
                cloblength = cloblength - cloblength
            Else
                tempvar = clob.Substring(tempcount, 3998)
                tempcount = tempcount + 3998
                cloblength = cloblength - 3998
            End If
            slno = slno + 1

            processmessage1(cloblength)
            'TRY
            sql = "insert into z_du_bulk (slno,bulkdata) values (" & slno & ",'" & tempvar & "')"
            Dim cmd152 As New OracleCommand(sql, conn)
            cmd152.ExecuteNonQuery()
            'Catch ex As Exception
            '    MsgBox(tempvar)
            'End Try

        Loop
        processmessage1("")
        processmessage("Formatting data - " & filename)

        sql = "PKGMISTOOL2.BULKUPLOAD_WITHOUT_TRIM"
        Dim cmd1 As New OracleCommand(sql, conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.Add("FILENAME", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = filename
        cmd1.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()

    End Sub

    Sub uploadfiledata_without_trim_MigrationTool(ByVal filename As String, ByVal username As String, ByVal du_clear_flag As String)

        Dim clob As String
        Dim cloblength As Integer
        Dim slno As Integer

        Try

            clob = My.Computer.FileSystem.ReadAllText(filename)
            clob = clob.Replace("'", "`")
            cloblength = clob.Length

        Catch ex As Exception

            clob = ""
            cloblength = 0
            oracle_execute_non_query("ten", username, username, "INSERT INTO C_MISPRINT (SERIALNO,REPORTDATA) VALUES (9999999999,'Visual Studio - Error in reading file " & filename & "')")

        End Try

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        If du_clear_flag = "Y" Then

            processmessage("Deleting existing data")

            oracle_execute_non_query("ten", username, username, "truncate table z_du")

        End If

        processmessage("Uploading data - " & filename)

        oracle_execute_non_query("ten", username, username, "truncate table z_du_bulk")
        'oracle_execute_non_query("ten", username, username, "truncate table z_email")

        tempcount = 0
        slno = 0

        Do Until cloblength = 0

            If cloblength < 3998 Then
                tempvar = clob.Substring(tempcount, cloblength)
                tempcount = tempcount + cloblength
                cloblength = cloblength - cloblength
            Else
                tempvar = clob.Substring(tempcount, 3998)
                tempcount = tempcount + 3998
                cloblength = cloblength - 3998
            End If
            slno = slno + 1

            processmessage1(cloblength)
            'TRY
            sql = "insert into z_du_bulk (slno,bulkdata) values (" & slno & ",'" & tempvar & "')"
            Dim cmd152 As New OracleCommand(sql, conn)
            cmd152.ExecuteNonQuery()
            'Catch ex As Exception
            '    MsgBox(tempvar)
            'End Try

        Loop
        processmessage1("")
        processmessage("Formatting data - " & filename)

        sql = "PKGMISTOOL2.BULKUPLOAD_WITHOUT_TRIM"
        Dim cmd1 As New OracleCommand(sql, conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.Add("FILENAME", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = filename
        cmd1.ExecuteNonQuery()


        conn.Close()
        conn.Dispose()

    End Sub

    Sub sendemail(ByVal sendfromaccount As String, ByVal database As String, ByVal user As String, ByVal password As String)

        ''Generating EMail
        ''Add "Microsoft Outlook 15.0 Object Library" in Project >> Reference >> Com
        ''Add the following in declaration part
        ''Imports System.Runtime.InteropServices
        ''Imports Outlook = Microsoft.Office.Interop.Outlook

        Dim oradb As String = "Data Source=" & database & ";User Id= " & user & ";Password= " & password & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Sending Mail")

        Dim sql1 As String
        Dim sql2 As String
        Dim tmplength As Integer = 0
        Dim substringno As Integer = 0
        Dim substringstart As Integer = 1
        Dim Bodydata As String = ""

        tempcount = 0
        sql = "SELECT MAIL_DATAID,MAIL_DATASUBID,MAIL_TO,NVL(MAIL_CC,'A')MAIL_CC,NVL(MAIL_BCC,'A')MAIL_BCC,MAIL_SUBJECT FROM Z_EMAIL ORDER BY MAIL_DATAID, MAIL_DATASUBID"
        Dim cmd21 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd21.ExecuteReader()
        While dr.Read()
            Bodydata = ""
            tmplength = 0
            substringno = 0
            substringstart = 1

            sql1 = "SELECT LENGTH(MAIL_BODY) TEMPLENGTH FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
            Dim cmd22 As New OracleCommand(sql1, conn)
            Dim dr1 As OracleDataReader = cmd22.ExecuteReader()
            While dr1.Read()
                tmplength = dr1.Item("TEMPLENGTH")
                Do Until tmplength = 0

                    If tmplength >= 3000 Then

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & ",3000) ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()

                    Else

                        substringno = tmplength

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & "," & substringno & ") ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()
                        tmplength = 0

                    End If

                Loop

            End While
            dr1.Close()

            Dim oApp As Outlook._Application
            oApp = New Outlook.Application()
            Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

            Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, sendfromaccount)
            newMail.To = dr.Item("MAIL_TO")
            If dr.Item("MAIL_TO") = "Error" Then
                MsgBox("Error Occured. Please check the table Z_EMAIL.", MsgBoxStyle.Critical, "File Missing")
                Exit Sub
            End If
            If dr.Item("MAIL_CC") <> "A" Then
                newMail.CC = dr.Item("MAIL_CC")
            End If
            If dr.Item("MAIL_BCC") <> "A" Then
                newMail.BCC = dr.Item("MAIL_BCC")
            End If
            newMail.Subject = dr.Item("MAIL_SUBJECT")
            newMail.HTMLBody = Bodydata
            newMail.SendUsingAccount = account
            newMail.Send()
            tempcount = tempcount + 1
            processmessage("Sending Mail No - " & tempcount)

        End While
        dr.Close()

        conn.Close()
        conn.Dispose()

        processmessage("")

        MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Sub checkfile(ByVal filename As String, ByVal message As String)

        If Not File.Exists(filename) Then

            MsgBox(message, MsgBoxStyle.Critical, "File Missing")
            Exit Sub

        End If

    End Sub

    Function checkaccountfile(ByVal filepath As String, ByVal datetocheck As Date) As Integer
        checkaccountfile = 0
        Dim datestring As String = datetocheck.ToString.Replace("/", "").Substring(0, 8)
        For Each file1 As String In Directory.GetDirectories("C:/du")
            Dim folername_date As String
            folername_date = Path.GetFileName(file1).Substring(0, 8)

            If folername_date = datestring Then
                If File.Exists(Path.GetFullPath(file1) & "\" & "DEP_Shadow_file.txt.gz") Then
                    Dim fi As New FileInfo(Path.GetFullPath(file1) & "\" & "DEP_Shadow_file.txt.gz")

                    Using inFile As FileStream = fi.OpenRead()
                        ' Get orignial file extension, for example "doc" from report.doc.gz. 
                        Dim curFile As String = fi.FullName
                        Dim origName = curFile.Remove(curFile.Length - fi.Extension.Length)

                        ' Create the decompressed file. 
                        Using outFile As FileStream = File.Create(origName)
                            Using Decompress As GZipStream = New GZipStream(inFile, CompressionMode.Decompress)
                                ' Copy the compressed file into the decompression stream. 
                                Dim buffer As Byte() = New Byte(4096) {}
                                Dim numRead As Integer
                                numRead = Decompress.Read(buffer, 0, buffer.Length)
                                Do While numRead <> 0
                                    outFile.Write(buffer, 0, numRead)
                                    numRead = Decompress.Read(buffer, 0, buffer.Length)
                                Loop
                                Console.WriteLine("Decompressed: {0}", fi.Name)

                            End Using
                        End Using
                    End Using

                    checkaccountfile = 1
                    Exit For
                    Exit Function
                End If
            End If
        Next

    End Function

    Sub oracle_execute_non_query(ByVal database As String, ByVal user As String, ByVal password As String, ByVal query As String)

        Dim oradb As String = "Data Source=" & database & ";User Id= " & user & ";Password= " & password & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim cmd1 As New OracleCommand(query, conn)
        cmd1.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()

    End Sub

    Sub processmessage(ByVal message As String)

        lblstatus.Text = message
        Application.DoEvents()

    End Sub

    Sub processmessage1(ByVal message As String)

        lblstatus2.Text = message
        Application.DoEvents()

    End Sub

    Private Function readNthLine(ByVal fileAndPath As String, ByVal lineNumber As Integer) As String
        Dim nthLine As String = Nothing
        Dim n As Integer
        Try

            Using sr As StreamReader = New StreamReader(fileAndPath)
                n = 0
                Do While (sr.Peek() >= 0) And (n < lineNumber)
                    sr.ReadLine()
                    n += 1
                Loop
                If sr.Peek() >= 0 Then
                    nthLine = sr.ReadLine()
                End If
            End Using
        Catch ex As Exception
            Throw
        End Try
        Return nthLine
    End Function

    'Public Sub CreateExcelFromCsvFile(ByVal strFolderPath As String, ByVal strFileName As String)
    '    Dim newFileName As String = "NewExcelFile.xls"
    '    Dim oExcelFile As Object
    '    ' Open Excel application object
    '    Try
    '        oExcelFile = GetObject(, "Excel.Application")
    '    Catch
    '        oExcelFile = CreateObject("Excel.Application")
    '    End Try
    '    oExcelFile.Visible = False
    '    oExcelFile.Workbooks.Open(strFolderPath + "\" + strFileName)
    '    ' Turn off message box so that we do not get any messages
    '    oExcelFile.DisplayAlerts = False
    '    ' Save the file as XLS file
    '    oExcelFile.ActiveWorkbook.SaveAs(Filename:=strFolderPath + "\" + newFileName, FileFormat:=Excel.XlFileFormat.xlExcel5, CreateBackup:=False)
    '    ' Close the workbook
    '    oExcelFile.ActiveWorkbook.Close(SaveChanges:=False)
    '    ' Turn the messages back on
    '    oExcelFile.DisplayAlerts = True
    '    ' Quit from Excel
    '    oExcelFile.Quit()
    '    ' Kill the variable
    '    oExcelFile = Nothing
    'End Sub

    Public Sub createdistrictbranchfolders(ByVal path)
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()
        sql = "SELECT DISTINCT UPPER(DTNAME) DTNAME FROM C_MISONLINEDATE ORDER BY DTNAME"
        Dim cmd As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd.ExecuteReader()
        While dr.Read()
            tempvar = path & "\" & dr.Item("DTNAME")
            System.IO.Directory.CreateDirectory(tempvar)
            sql = "SELECT SOLID2,UPPER(SOLNAME) SOLNAME FROM C_MISONLINEDATE WHERE DTNAME = '" & dr.Item("DTNAME") & "'"
            Dim cmd1 As New OracleCommand(sql, conn)
            Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
            While dr1.Read()
                tempvar = path & "\" & dr.Item("DTNAME") & "\" & dr1.Item("SOLID2") & "_" & dr1.Item("SOLNAME")
                If (Not System.IO.Directory.Exists(tempvar)) Then
                    System.IO.Directory.CreateDirectory(tempvar)
                End If
            End While
            dr1.Close()
        End While
        dr.Close()
        conn.Close()
        conn.Dispose()
    End Sub
    'Public Sub CreateExcelFromCsvFile(ByVal strFolderPath As String, ByVal strFileName As String, ByVal strfileextension As String)
    '    Dim newFileName As String = strFileName & ".xls"
    '    Dim oExcelFile As Object
    '    Try
    '        oExcelFile = GetObject(, "Excel.Application")
    '    Catch
    '        oExcelFile = CreateObject("Excel.Application")
    '    End Try
    '    oExcelFile.Visible = False
    '    oExcelFile.Workbooks.Open(strFolderPath + "\" + strFileName + "." + strfileextension)
    '    oExcelFile.DisplayAlerts = False
    '    oExcelFile.ActiveWorkbook.SaveAs(Filename:=strFolderPath + "\" + newFileName, FileFormat:=Excel.XlFileFormat.xlExcel5, CreateBackup:=False)
    '    oExcelFile.ActiveWorkbook.Close(SaveChanges:=False)
    '    oExcelFile.DisplayAlerts = True
    '    oExcelFile.Quit()
    '    oExcelFile = Nothing
    'End Sub

    Public Function GetFileName(ByVal filepath As String) As String
        'This Function Gets the name of a file without the path or extension.
        Dim slashindex As Integer = filepath.LastIndexOf("\")
        Dim dotindex As Integer = filepath.LastIndexOf(".")
        GetFileName = filepath.Substring(slashindex + 1, dotindex - slashindex - 1)
    End Function
    'Private Sub formatexcel(ByVal filename)
    '    Dim oExel As Excel.Application
    '    Dim oWorkbook As Excel.Workbook
    '    Dim oWorksheet As Excel.Worksheet
    '    Dim oRange As Excel.Range
    '    Dim rCnt As Integer
    '    Dim cCnt As Integer
    '    Dim Obj As Object
    '    Dim sReplace As String = "ABC"
    '    oExel = CreateObject("Excel.Application")
    '    oWorkbook = oExel.Application.Workbooks.Open(filename)
    '    oExel.Application.Interactive = True
    '    oExel.Application.UserControl = True
    '    For Each oWorksheet In oExel.ActiveWorkbook.Worksheets
    '        oRange = oWorksheet.UsedRange
    '        For rCnt = 1 To oRange.Rows.Count
    '            For cCnt = 1 To oRange.Columns.Count
    '                Obj = CType(oRange.Cells(rCnt, cCnt), Excel.Range).Text
    '                If Obj <> Nothing Then
    '                    ' find and replace
    '                    'MessageBox.Show(Obj)
    '                End If

    '            Next
    '        Next
    '    Next
    '    oWorkbook.Save()
    '    oWorkbook.Close()
    '    oExel.Quit()
    '    oExel = Nothing
    'End Sub

    'Sub option28()          'Generate Word file

    '    Dim solid As String
    '    Dim solname As String

    '    Dim lpdsuit As Integer
    '    Dim lpdrr As Integer
    '    Dim lpdothers As Integer

    '    Dim tot_npa As Integer
    '    Dim npa_without_action As Integer
    '    Dim npa_without_action_march As Integer

    '    Dim suit_pend As Integer
    '    Dim rr_pend As Integer
    '    Dim legal_entered As Integer


    '    Dim oWord As Word.Application
    '    Dim oDoc As Word.Document
    '    Dim oTable As Word.Table, oTable1 As Word.Table
    '    Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
    '    Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph, oPara5 As Word.Paragraph, oPara6 As Word.Paragraph
    '    Dim oRng As Word.Range
    '    Dim count As Integer

    '    count = 0


    '    Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
    '    Dim conn As New OracleConnection(oradb)
    '    conn.Open()

    '    'Start Word and open the document template.
    '    oWord = CreateObject("Word.Application")
    '    oWord.Visible = True
    '    oDoc = oWord.Documents.Add

    '    'Setting page margin
    '    oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.0)
    '    oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.0)
    '    oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.75)
    '    oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.75)

    '    'Justify
    '    oDoc.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify




    '    oDoc.Range.Font.Name = "Abadi MT Condensed Light"
    '    oDoc.Range.Font.Size = 5
    '    oDoc.Paragraphs.Style = "No Spacing"

    '    'Add a picture at the header
    '    'oDoc.Content.Application.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader
    '    'oDoc.Content.Application.Selection.Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\1.jpg")


    '    'Dim PIctureLocation As String = "E:\VBProject\1.jpg"  --->Defining picture location
    '    'oDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\1.jpg")


    '    'Add picture in footer
    '    ''oDoc.Content.Application.Selection.Fields.Add(Range:=oDoc.Content.Application.Selection.Range, Type:=CInt(Word.WdFieldType.wdFieldEmpty), Text:="page")
    '    'oDoc.Content.Application.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument
    '    'oDoc.Content.Application.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter
    '    'oDoc.Content.Application.Selection.Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\2.jpg")
    '    ''oDoc.Content.Application.Selection.TypeText(Text:="Martens")

    '    'return to the main document        
    '    ' oDoc.Content.Application.ActiveWindow.ActivePane.View.SeekView = CInt(Word.WdSeekView.wdSeekMainDocument)

    '    sql = "SELECT TEXT1 ,TEXT20,NUMBER1,NUMBER2,NUMBER3,NUMBER4,NUMBER5,NUMBER6,NUMBER7,NUMBER8,NUMBER9 FROM C_MISADV WHERE C_ACID = 'SOLNAME' ORDER BY TEXT1"
    '    Dim cmd4 As New OracleCommand(sql, conn)
    '    Dim dr As OracleDataReader = cmd4.ExecuteReader()


    '    While dr.Read()
    '        solid = dr("text1").ToString()
    '        solname = dr("text20").ToString()

    '        lpdsuit = dr("number1")
    '        lpdrr = dr("number2")
    '        lpdothers = dr("number3")

    '        tot_npa = dr("number4")
    '        npa_without_action = dr("number5")
    '        npa_without_action_march = dr("number6")

    '        suit_pend = dr("number7")
    '        rr_pend = dr("number8")
    '        legal_entered = dr("number9")

    '        'Inserting a picture file
    '        oDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\1.jpg")

    '        oPara1 = oDoc.Content.Paragraphs.Add
    '        oPara1.Format.SpaceAfter = 5
    '        oPara1.Range.InsertParagraphAfter()

    '        oPara1 = oDoc.Content.Paragraphs.Add()
    '        oPara1.Range.Text = "The Branch Manager"
    '        oPara1.Format.SpaceAfter = 1
    '        oPara1.Style = "No Spacing"
    '        oPara1.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
    '        oPara1.Range.Font.Size = 10
    '        oPara1.Range.InsertParagraphAfter()
    '        oPara1.Range.Text = "Kerala Gramin Bank"
    '        oPara1.Format.SpaceAfter = 1
    '        oPara1.Range.InsertParagraphAfter()
    '        oPara1.Range.Text = solname
    '        oPara1.Format.SpaceAfter = 5  'Setting space.
    '        oPara1.Range.InsertParagraphAfter()

    '        'Insert a paragraph at the end of the document.
    '        '** \endofdoc is a predefined bookmark.
    '        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        'oPara2.Format.SpaceAfter = 25
    '        oPara2.Range.Text = "Sir,"
    '        oPara2.Style = "No Spacing"
    '        oPara2.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
    '        oPara2.Range.Font.Size = 10

    '        oPara2.Format.SpaceAfter = 5
    '        oPara2.Range.InsertParagraphAfter()

    '        oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara3.Range.Text = "Sub: NPA Accounts with no action"
    '        oPara3.Style = "No Spacing"
    '        oPara3.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
    '        oPara3.Range.Font.Size = 10
    '        oPara3.Format.SpaceAfter = 5
    '        oPara3.Range.InsertParagraphAfter()
    '        oPara3.Style = "No Spacing"
    '        oPara3.Range.Text = "Furnished here below are the action initiated accounts (LPD Suit, LPD RR, LPD others), total number of NPA accounts and the number of accounts marked as NPA before 01/04/2013 and lying without any recovery action."

    '        oPara3.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
    '        oPara3.Range.Font.Size = 10
    '        oPara3.Format.SpaceAfter = 1
    '        oPara3.Range.InsertParagraphAfter()

    '        oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara3.SpaceAfter = 2

    '        'Create a table with 8 rows and 2 columns
    '        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 2)
    '        oTable.Range.ParagraphFormat.SpaceAfter = 2

    '        For r = 1 To 8
    '            For c = 1 To 2

    '                If r = 1 Then
    '                    oTable.Cell(r, c).Range.Font.Bold = True


    '                    If c = 1 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

    '                        oTable.Cell(r, c).Range.Text = "Head"
    '                    End If

    '                    If c = 2 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = "No. of A\c"
    '                    End If

    '                ElseIf r = 2 Then

    '                    If c = 1 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable.Cell(r, c).Range.Text = "LPD Suit"
    '                    Else
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable.Cell(r, c).Range.Text = lpdsuit
    '                    End If

    '                ElseIf r = 3 Then

    '                    If c = 1 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = "LPD RR"
    '                    Else
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = lpdrr
    '                    End If


    '                ElseIf r = 4 Then

    '                    If c = 1 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = "LPD Others"

    '                    Else
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = lpdothers
    '                    End If



    '                ElseIf r = 5 Then
    '                    If c = 1 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = "Total LPD"
    '                    Else
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = lpdothers + lpdrr + lpdsuit
    '                    End If



    '                ElseIf r = 6 Then
    '                    If c = 1 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = "Total NPA"
    '                    Else
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = tot_npa
    '                    End If


    '                ElseIf r = 7 Then

    '                    If c = 1 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = "Of which, NPA Accounts lying without action"
    '                    Else
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = npa_without_action
    '                    End If


    '                ElseIf r = 8 Then
    '                    If c = 1 Then
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable.Cell(r, c).Range.Text = "                 NPA Accounts marked before March 2013 lying without action"
    '                    Else
    '                        oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter

    '                        oTable.Cell(r, c).Range.Text = npa_without_action_march
    '                    End If
    '                Else
    '                    oTable.Cell(r, c).Range.Text = "r" & r & "c" & c
    '                End If
    '                oTable.Cell(r, c).Borders.Enable = True
    '            Next
    '        Next
    '        oTable.Columns.Item(1).Width = oWord.InchesToPoints(5.7)   'Change width of columns 1 & 2
    '        oTable.Columns.Item(2).Width = oWord.InchesToPoints(1.1)

    '        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara4.Format.SpaceAfter = 2

    '        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara4.Format.SpaceAfter = 5

    '        oPara4.Range.Text = "Data entry status of LPD accounts under LPD module as instructed vide circular 3/R&L/2013 dated 17/07/2013 is given below:"
    '        oPara4.Style = "No Spacing"
    '        oPara4.Range.Font.Name = "Calibri (Body)"
    '        oPara4.Range.Font.Size = 10

    '        '--To get bold and underline
    '        'oPara4.Range.Font.Bold = True
    '        'oPara4.Range.Font.Underline = True

    '        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara4.Format.SpaceAfter = 2
    '        oPara4.Range.InsertParagraphAfter()

    '        oPara4.Range.Font.Bold = False
    '        oPara4.Range.Font.Underline = False

    '        oTable1 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 4)
    '        oTable1.Range.ParagraphFormat.SpaceAfter = 2

    '        For r = 1 To 5
    '            For c = 1 To 4

    '                If r = 1 Then

    '                Else
    '                    oTable1.Cell(r, c).Borders.Enable = True
    '                End If

    '                If r = 2 Then

    '                    If c = 1 Then
    '                        oTable1.Cell(r, c).Range.Font.Bold = True
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "Module"
    '                    ElseIf c = 2 Then
    '                        oTable1.Cell(r, c).Range.Font.Bold = True
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "Total"
    '                    ElseIf c = 3 Then
    '                        oTable1.Cell(r, c).Range.Font.Bold = True
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "Entered"
    '                    Else
    '                        oTable1.Cell(r, c).Range.Font.Bold = True
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "Pending"
    '                    End If

    '                ElseIf r = 3 Then

    '                    If c = 1 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "Suit"
    '                    ElseIf c = 2 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = lpdsuit
    '                    ElseIf c = 3 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = lpdsuit - suit_pend
    '                    Else
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = suit_pend
    '                    End If

    '                ElseIf r = 4 Then

    '                    If c = 1 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "RR"

    '                    ElseIf c = 2 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = lpdrr
    '                    ElseIf c = 3 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = lpdrr - rr_pend
    '                    Else
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = rr_pend
    '                    End If

    '                ElseIf r = 5 Then

    '                    If c = 1 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "Legal action waived"

    '                    ElseIf c = 2 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "XX"
    '                    ElseIf c = 3 Then
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = legal_entered
    '                    Else
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
    '                        oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
    '                        oTable1.Cell(r, c).Range.Text = "XX"
    '                    End If
    '                End If

    '            Next

    '        Next

    '        oTable1.Columns.Item(1).Width = oWord.InchesToPoints(1.7)   'Change width of columns 1 & 2
    '        oTable1.Columns.Item(2).Width = oWord.InchesToPoints(1.7)
    '        oTable1.Columns.Item(3).Width = oWord.InchesToPoints(1.7)
    '        oTable1.Columns.Item(4).Width = oWord.InchesToPoints(1.7)

    '        oTable1.Cell(1, 1).Merge(MergeTo:=oTable1.Cell(1, 4))
    '        oTable1.Cell(1, 1).Range.Font.Bold = True
    '        oTable1.Cell(1, 1).Range.Text = "Data entry status in LPD Module"

    '        oTable1.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
    '        oTable1.Cell(1, 1).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter

    '        oTable1.Cell(1, 1).Borders.Enable = True


    '        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara4.Format.SpaceAfter = 2
    '        oPara4.Range.ParagraphFormat.Alignment = 3
    '        oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara4.Range.Text = "Branch is advised to make detailed study of the above data pertaining to NPA, LPD, Non LPD and LPD module and take immediate steps as here under:"
    '        oPara4.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
    '        oPara4.Range.Font.Size = 10
    '        oPara4.Format.SpaceAfter = 5
    '        oPara4.Range.InsertParagraphAfter()


    '        oPara4.Range.ParagraphFormat.Alignment = 3
    '        oPara4.Format.SpaceAfter = 2
    '        oPara4.Range.ListFormat.ApplyBulletDefault() 'Bullet

    '        oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara5.Format.SpaceAfter = 5
    '        oPara5.Range.ParagraphFormat.Alignment = 3

    '        oPara5.Range.Text = "Generate a statement by accessing NPARPT 411 and get the list of accounts which were marked as NPA prior to 01/04/2013 and is remaining without any action. Follow up each of these accounts and ensure recovery of full overdue/regularization/closure/action before 28/02/2014."
    '        oPara5.Range.InsertParagraphAfter()

    '        oPara6 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara6.Range.Text = "Verify all LPD accounts and complete the work, relating to entering the data of suit filed accounts, RR initiated accounts, Legal action waived accounts in the system by accessing the menu Suit / RR / LAW."
    '        oPara6.Format.SpaceAfter = 2

    '        oPara6.Range.InsertParagraphAfter()

    '        oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara5.Range.Text = "Updation of LPD module is very urgent for follow up and data generation purposes. Hence the work should be completed on a war footing basis before 15/02/2014."
    '        oPara5.Format.SpaceAfter = 2

    '        oPara5.Range.InsertParagraphAfter()

    '        oPara5.Range.Text = "A confirmation letter regarding completion of the above actions to be submitted to concerned RO by 01/03/2014."
    '        oPara5.Format.SpaceAfter = 75

    '        oPara5.Range.InsertParagraphAfter()
    '        oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
    '        oPara5.Range.Font.Bold = True

    '        oPara4.Range.ListFormat.RemoveNumbers()

    '        oPara5.Range.Text = "S.Radhakrishnan Nair"
    '        oPara5.Format.SpaceAfter = 2
    '        oPara5.Range.InsertParagraphAfter()
    '        oPara5.Range.Text = "General Manager"
    '        oPara5.Format.SpaceAfter = 60

    '        oPara5.Range.InsertParagraphAfter()

    '        'Page break
    '        oDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\2.jpg")
    '        oRng = oDoc.Bookmarks.Item("\endofdoc").Range
    '        oRng.ParagraphFormat.SpaceAfter = 1
    '        oRng.InsertBreak(Word.WdBreakType.wdPageBreak)
    '        oRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

    '        count = count + 1
    '    End While
    '    dr.Close()


    '    MsgBox("Generated " & count & " pages", MsgBoxStyle.Information, "Invalid date")

    'End Sub

    Sub option29()          'Mobile Banking Transaction Status

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "Y")

                End If

            Next

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - UPLOAD_KGB_MOB_TRAN")

            sql = "PKGEMAIL110.UPLOAD_KGB_MOB_TRAN"
            Dim cmd1 As New OracleCommand(sql, conn)
            cmd1.CommandType = CommandType.StoredProcedure
            cmd1.ExecuteNonQuery()


            processmessage("Package - DATAID_1102")

            sql = "PKGEMAIL110.DATAID_1102"
            Dim cmd2 As New OracleCommand(sql, conn)
            cmd2.CommandType = CommandType.StoredProcedure
            cmd2.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
            cmd2.ExecuteNonQuery()

            'sendemail("smgbmis3@gmail.com", "ten", username, username)
            'sendemail("mis@kgbmis.in", "ten", username, username)
            sendemail("mis@kgbmis.in", "ten", username, username)
            'processmessage("Package - finished")

        End If

    End Sub

    Sub option30()          'Create A folder in Specified path

        Dim foldername As String
        Dim folderpath As String
        Dim solid As String
        Dim solname As String
        Dim dtname As String
        Dim roname As String

        Dim count As Integer

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        count = 0

        folderpath = InputBox("Enter the path", "Enter value", "C:\du")

        If folderpath = "" Then
            MsgBox("Enter path")
        End If

        'Handling issues like entering C or D  drive for path
        If (folderpath.Length = 1) Then
            folderpath = folderpath & ":\"
        End If

        If folderpath(folderpath.Length - 1) = "\" Or folderpath(folderpath.Length - 1) = "/" Then
            folderpath = folderpath.Remove(folderpath.Length - 1)
        End If

        'Handling entered path in D\abc

        If folderpath(1) <> ":" Then
            folderpath = folderpath.Insert(1, ":")
        End If

        folderpath = folderpath.Replace("/", "\")

        foldername = InputBox("Enter Bank <S>MGB; <N>MGB; <R>O; <D>ISTRICT; <K>GB;", "", "K")
        If foldername = "K" Or foldername = "k" Then
            sql = "SELECT KGB_SOLID,KGB_SOLNAME,KGB_DISTRICT,KGB_RO FROM Z_KGB"
        ElseIf foldername = "S" Or foldername = "s" Then
            sql = "SELECT KGB_SOLID,KGB_SOLNAME,KGB_DISTRICT,KGB_RO FROM Z_KGB WHERE KGB_SOLID >40000 "
        ElseIf foldername = "n" Or foldername = "N" Then
            sql = "SELECT KGB_SOLID,KGB_SOLNAME,KGB_DISTRICT,KGB_RO FROM Z_KGB WHERE KGB_SOLID <40000 "
        ElseIf foldername = "r" Or foldername = "R" Then
            sql = "SELECT DISTINCT KGB_RO FROM Z_KGB "
        ElseIf foldername = "d" Or foldername = "D" Then
            sql = "SELECT DISTINCT KGB_DISTRICT FROM Z_KGB "
        Else
            'cnn.Close()
            Exit Sub
        End If


        Dim cmd4 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd4.ExecuteReader()

        solid = ""
        solname = ""
        dtname = ""
        roname = ""

        If foldername.ToUpper() = "D" Then
            While dr.Read()
                count = count + 1
                dtname = dr("KGB_DISTRICT")
                tempvar = folderpath & "\" & dtname

                'If directory already exist delete it with its contents
                If Directory.Exists(tempvar) Then
                    System.IO.Directory.Delete(tempvar, True)
                End If
                System.IO.Directory.CreateDirectory(tempvar)
            End While
            dr.Close()

        ElseIf foldername.ToUpper() = "R" Then
            While dr.Read()
                count = count + 1
                roname = dr("KGB_RO")
                tempvar = folderpath & "\" & roname
                If Directory.Exists(tempvar) Then
                    System.IO.Directory.Delete(tempvar, True)
                End If
                System.IO.Directory.CreateDirectory(tempvar)
            End While
            dr.Close()
        Else
            While dr.Read()
                count = count + 1
                solid = dr("KGB_SOLID")
                solname = dr("KGB_SOLNAME")
                dtname = dr("KGB_DISTRICT")
                roname = dr("KGB_RO")
                solid = solid.ToString.PadLeft(5, "0") 'paadding 0 for eNMGB branches to adjust 5 digit code
                solname = solname.ToString().Replace(":", " ")
                tempvar = folderpath & "\" & solid & "_" & dtname & "_" & roname & "_" & solname

                If Directory.Exists(tempvar) Then
                    System.IO.Directory.Delete(tempvar, True)
                End If
                System.IO.Directory.CreateDirectory(tempvar)
            End While
            dr.Close()
        End If

        cnn.Close()
        MsgBox(count & "Folders  created. Process completed")

    End Sub

    Sub option31()                    'File copy
        Dim sourcefile As String
        Dim includesubfolder As String
        Dim copy_tofolder As String

        sourcefile = InputBox("Enter the source file to be copied with full path")

        If File.Exists(sourcefile) = False Then
            MsgBox("Cannot find the source file. Please check", MsgBoxStyle.Critical)
            Exit Sub
        End If

        copy_tofolder = InputBox("Enter the path to which the file to be copied", "Copy to", "C:\du")
        includesubfolder = InputBox("Include sub folder (Y/N)", "", "Y")

        'Create folder if no destination folder exist
        If Directory.Exists(copy_tofolder) = False Then
            Directory.CreateDirectory(copy_tofolder)
        End If

        'Copy file
        If includesubfolder.ToUpper() = "N" Then
            copy_tofolder = Path.Combine(copy_tofolder, Path.GetFileName(sourcefile))
            File.Copy(sourcefile, copy_tofolder)
        Else

            'Handling for subfolders  -Copy file function is recursive  second and third argument will be changed for sub folders
            For Each dir1 As String In Directory.GetDirectories(copy_tofolder)

                Dim temp As String
                temp = ""
                temp = copy_tofolder
                temp = Path.Combine(copy_tofolder, Path.GetFileName(sourcefile))
                copyfile(sourcefile, copy_tofolder, dir1)
                File.Copy(sourcefile, temp, True)
            Next

        End If
        MsgBox("Process completed", MsgBoxStyle.Information)
    End Sub

    Sub option32()   'Execute Script file
        Dim scriptfilename As String
        Dim script_exepath As String
        Dim script_exedb As String
        Dim includesub As String


        scriptfilename = InputBox("Script file name (with path)")
        If File.Exists(scriptfilename) = False Then
            MsgBox("Cannot find the Source file. Please check", MsgBoxStyle.Critical)
            Exit Sub
        End If

        script_exedb = InputBox("Access database file name (wihout path)", "Database Name", "NMGB.mdb")
        script_exepath = InputBox("Access database file path (without file name)", "Database Path", "C:\du")
        script_exedb = "Server\" & script_exedb


        includesub = InputBox("Update subfolders   - Y/N", "Enter value", "Y")

        'Executing Script in Single folder file
        If includesub.ToUpper = "N" Then
            Dim filevar As String = scriptfilename
            Dim Line As String = "A"

            If File.Exists(script_exepath & "\" & script_exedb) = False Then
                MsgBox("Cannot find the Destination file. Please check", MsgBoxStyle.Critical)
                Exit Sub
            End If

            Dim cnn As New OleDb.OleDbConnection
            cnn = New OleDb.OleDbConnection

            Dim strConnection As String
            strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & script_exepath & "\" & script_exedb
            cnn.ConnectionString = strConnection

            Try
                If Not cnn.State = ConnectionState.Open Then
                    cnn.Open()
                End If
            Catch ex As Exception
                MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
            End Try

            'Reading file content
            If System.IO.File.Exists(filevar) = True Then
                Dim objReader As New System.IO.StreamReader(filevar)
                Do While objReader.Peek() <> -1
                    Line = Line & objReader.ReadLine() & vbNewLine
                    'Line = Line.Remove(0, 2)

                    Line = readNthLine(scriptfilename, 0)

                    Try
                        Dim cmd As New OleDb.OleDbCommand
                        cmd.CommandText = Line
                        cmd.Connection = cnn
                        cmd.ExecuteNonQuery()  'Executing command
                    Catch ex As Exception
                        MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                        'cnn.Close()
                        'Exit Sub
                    End Try
                Loop
                objReader.Close()
            Else
                MsgBox("File Does Not Exist")

            End If
        ElseIf includesub.ToUpper() = "Y" Then

            'Executing Script in sub folders recursive function second and third argument will be changed for sub folders
            executescriptInsubfolder(scriptfilename, script_exepath, script_exedb)
            processmessage("Executing script in file " & script_exepath & "\" & script_exedb)
        Else
            MsgBox("Enter either Y or N")
            Exit Sub

        End If

        MsgBox("Process completed", MsgBoxStyle.Information)

    End Sub
    Sub option33()   'Basedata Generation Timing Email

        ' Checking whether email3.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\email3.txt"

        checkfile(file1, "Rename the EMail file 40103_XX-XX-XXXX.email as email3.txt and place in c:/du folder")

        uploadfiledata(file1, username, "Y")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        processmessage("Package - DATAID_1103")

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGEMAIL110.DATAID_1103"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, ParameterDirection.Input).Value = "ALL"
        cmd5.ExecuteNonQuery()

        sendemail("mis@kgbmis.in", "ten", username, username)

    End Sub

    Sub option34()      'STAFF Upload file creation

        ' Checking whether  files exists

        Dim tempvar As String
        Dim tempcount As String = 0

        processmessage("Checking files")

        file1 = "c:\du\STAFF_NAME.TXT"
        file2 = "c:\du\STAFF_BM.TXT"

        checkfile(file1, "Place the File naming as STAFF_NAME.TXT")
        checkfile(file2, "Place the File naming as STAFF_BM.TXT")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")

        ' Connecting to oracle data base

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        ' Calling packages

        processmessage("Package - PKGMISTOOL2.STAFF_UPLOAD")

        sql = "PKGMISTOOL2.STAFF_UPLOAD"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Creating staff_name_upload.txt")

        tempvar = ""
        Dim sw1 As StreamWriter = New StreamWriter("c:/du/staff_name_upload.txt.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO < 5 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd12 As New OracleCommand(sql, conn)
        Dim dr1 As OracleDataReader = cmd12.ExecuteReader()
        While dr1.Read()
            tempvar = dr1.Item("REPORTDATA")
            sw1.WriteLine(tempvar)
        End While
        dr1.Close()
        sw1.Close()

        processmessage("Creating staff_bm_upload.txt")

        tempvar = ""
        Dim sw2 As StreamWriter = New StreamWriter("c:/du/staff_bm_upload.txt.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO BETWEEN 5 AND 8 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd13 As New OracleCommand(sql, conn)
        Dim dr2 As OracleDataReader = cmd13.ExecuteReader()
        While dr2.Read()
            tempvar = dr2.Item("REPORTDATA")
            sw2.WriteLine(tempvar)
        End While
        dr2.Close()
        sw2.Close()

        processmessage("Creating staff_pos_upload.txt")

        tempvar = ""
        Dim sw3 As StreamWriter = New StreamWriter("c:/du/staff_pos_upload.txt.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO >8 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd14 As New OracleCommand(sql, conn)
        Dim dr3 As OracleDataReader = cmd14.ExecuteReader()
        While dr3.Read()
            tempvar = dr3.Item("REPORTDATA")
            sw3.WriteLine(tempvar)
        End While
        dr3.Close()
        sw3.Close()

        processmessage("")

        MsgBox("Upload file created successfully", MsgBoxStyle.Information, "Process Completed")

        conn.Close()
        conn.Dispose()

    End Sub
    Sub option35()   'RO Follow Up Status

        processmessage("Checking files")

        file1 = "c:\du\email_old.txt"
        file2 = "c:\du\email_new.txt"
        'file3 = "c:\du\email_2.txt"

        checkfile(file1, "Rename one month back email file 40101_XX-XX-XXXX.email as email_old.txt and place in c:/du folder")
        checkfile(file2, "Rename previousday email file 40101_XX-XX-XXXX.email as email_new.txt and place in c:/du folder")
        'checkfile(file3, "Rename previousday second email file as email_2.txt and place in c:/du folder")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")
        'uploadfiledata(file3, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        processmessage("Package - DATAID_1121")

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGEMAIL112.DATAID_1121"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, ParameterDirection.Input).Value = "ALL"
        cmd5.ExecuteNonQuery()

        sendemail("smgbmis3@gmail.com", "ten", username, username)

    End Sub
    Sub option36()          'ATM Transaction Status 
        ''preetha
        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0
        Dim ason As Date

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else
            'oracle_execute_non_query("ten", username, username, "DROP TABLE C_MISPRINT")
            'oracle_execute_non_query("ten", username, username, "DROP TABLE Z_DU")
            'oracle_execute_non_query("ten", username, username, "DROP TABLE Z_DU_BULK")
            'oracle_execute_non_query("ten", username, username, "DROP TABLE Z_EMAIL")
            'oracle_execute_non_query("ten", username, username, "CREATE TABLE C_MISPRINT (SOLID  VARCHAR2(10),CUSTOMERID VARCHAR2(9),ACCOUNTNUMBER  VARCHAR2(16),SERIALNO NUMBER(30,10),SUBSERIALNO NUMBER(30,10),REPORTDATA VARCHAR2(4000))")
            'oracle_execute_non_query("ten", username, username, "CREATE INDEX C_MISPRINT_IDX ON C_MISPRINT (SERIALNO, SUBSERIALNO, ACCOUNTNUMBER, CUSTOMERID)")
            'oracle_execute_non_query("ten", username, username, "CREATE TABLE Z_DU (FILENAME VARCHAR2(100), LINENO NUMBER(10), LINEDATA VARCHAR2(4000), DATE1 DATE, DATE2 DATE, DATE3 DATE, DATE4 DATE, DATE5 DATE, DATE6 DATE, DATE7 DATE, DATE8 DATE, DATE9 DATE, DATE10 DATE, DATE11 DATE, DATE12 DATE, DATE13 DATE, DATE14 DATE, DATE15 DATE, DATE16 DATE, DATE17 DATE, DATE18 DATE, DATE19 DATE, DATE20 DATE, NUMBER1 NUMBER(20,2), NUMBER2 NUMBER(20,2), NUMBER3 NUMBER(20,2), NUMBER4 NUMBER(20,2), NUMBER5 NUMBER(20,2), NUMBER6 NUMBER(20,2), NUMBER7 NUMBER(20,2), NUMBER8 NUMBER(20,2), NUMBER9 NUMBER(20,2), NUMBER10 NUMBER(20,2), NUMBER11 NUMBER(20,2), NUMBER12 NUMBER(20,2), NUMBER13 NUMBER(20,2), NUMBER14 NUMBER(20,2), NUMBER15 NUMBER(20,2), NUMBER16 NUMBER(20,2), NUMBER17 NUMBER(20,2), NUMBER18 NUMBER(20,2), NUMBER19 NUMBER(20,2), NUMBER20 NUMBER(20,2), TEXT1 VARCHAR2(100), TEXT2 VARCHAR2(100), TEXT3 VARCHAR2(100), TEXT4 VARCHAR2(100), TEXT5 VARCHAR2(100), TEXT6 VARCHAR2(100), TEXT7 VARCHAR2(100), TEXT8 VARCHAR2(100), TEXT9 VARCHAR2(100), TEXT10 VARCHAR2(100), TEXT11 VARCHAR2(100), TEXT12 VARCHAR2(100), TEXT13 VARCHAR2(100), TEXT14 VARCHAR2(100), TEXT15 VARCHAR2(100), TEXT16 VARCHAR2(100), TEXT17 VARCHAR2(100), TEXT18 VARCHAR2(100), TEXT19 VARCHAR2(100), TEXT20 VARCHAR2(100))")
            'oracle_execute_non_query("ten", username, username, "CREATE INDEX Z_DU_IDX1 ON Z_DU (FILENAME, LINENO)")
            'oracle_execute_non_query("ten", username, username, "CREATE TABLE Z_DU_BULK (SLNO NUMBER(10),BULKDATA VARCHAR2(4000))")
            'oracle_execute_non_query("ten", username, username, "CREATE TABLE Z_EMAIL (MAIL_DATAID NUMBER(5),MAIL_TO VARCHAR2(4000),MAIL_CC VARCHAR2(4000),MAIL_BCC VARCHAR2(4000),MAIL_SUBJECT VARCHAR2(4000),MAIL_BODY CLOB,MAIL_DATASUBID VARCHAR2(10))")
            'oracle_execute_non_query("ten", username, username, "CREATE INDEX C_EMAIL_IDX1 ON Z_EMAIL (MAIL_DATAID, MAIL_DATASUBID)")
            'For Each dir In dirs

            '    tempcount = tempcount + 1

            '    If tempcount = 1 Then

            '        uploadfiledata_without_trim(dir, username, "Y")

            '    Else

            '        uploadfiledata_without_trim(dir, username, "N")

            '    End If

            'Next

            processmessage("Validating data")

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()
            sql = "SELECT  PKGEMAIL111.VALIDATE_FILES_BEFORE_UPDATE AA FROM DUAL"
            Dim cmd44 As New OracleCommand(sql, conn)
            Dim dr As OracleDataReader = cmd44.ExecuteReader()
            While dr.Read()
                Dim goutput As String = dr.Item("AA")
                If goutput.Substring(0, 1) = "9" Then
                    MsgBox(goutput.Substring(1, 99), MsgBoxStyle.Critical, "Error")
                    Exit Sub
                Else
                    ason = goutput.Substring(1, 10)
                End If
            End While
            dr.Close()
            'conn.Close()
            'conn.Dispose()

            ' Delete existing data, if any, from c_du table

            processmessage("Deleting existing data")

            oracle_execute_non_query("ten", username, username, "truncate table z_email")

            ' Calling packages

            processmessage("Package - UPLOAD_KGB_ATM_TRAN")

            sql = "PKGEMAIL111.UPLOAD_KGB_ATM_TRAN"
            Dim cmd1 As New OracleCommand(sql, conn)
            cmd1.CommandType = CommandType.StoredProcedure
            cmd1.ExecuteNonQuery()

            processmessage("Package - DATAID_1111")

            sql = "PKGEMAIL111.DATAID_1111"
            Dim cmd2 As New OracleCommand(sql, conn)
            cmd2.CommandType = CommandType.StoredProcedure
            cmd2.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = ason
            cmd2.ExecuteNonQuery()

            processmessage("Package - DATAID_1112")

            sql = "PKGEMAIL111.DATAID_1112"
            Dim cmd3 As New OracleCommand(sql, conn)
            cmd3.CommandType = CommandType.StoredProcedure
            cmd3.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = ason
            cmd3.ExecuteNonQuery()

            processmessage("Package - DATAID_1113")

            sql = "PKGEMAIL111.DATAID_1113"
            Dim cmd4 As New OracleCommand(sql, conn)
            cmd4.CommandType = CommandType.StoredProcedure
            cmd4.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = ason
            cmd4.ExecuteNonQuery()

            processmessage("Package - DATAID_1115")

            sql = "PKGEMAIL111.DATAID_1115"
            Dim cmd5 As New OracleCommand(sql, conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = ason
            cmd5.ExecuteNonQuery()

            'processmessage("Package - ATM_DASHBOARD")

            'sql = "PKGEMAIL111.ATM_DASHBOARD"
            'Dim cmd6 As New OracleCommand(sql, conn)
            'cmd6.CommandType = CommandType.StoredProcedure
            'cmd6.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = ason
            'cmd6.ExecuteNonQuery()

            'sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'ATM_ID'"
            'display_in_File(sql, "C:\du\AIDM.atm")
            'Process.Start("C:\du\AIDM.atm")

            'sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'NO_HIT'"
            'display_in_File(sql, "C:\du\NOHT.atm")
            'Process.Start("C:\du\NOHT.atm")

            'sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'ATM'"
            'display_in_File(sql, "C:\du\AHT.atm")
            'Process.Start("C:\du\AHT.atm")

            'sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'ATMCRD'"
            'display_in_File(sql, "C:\du\CHT.atm")
            'Process.Start("C:\du\CHT.atm")

            'sendemail("mis@kgbmis.in", "ten", username, username)

        End If
        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Info")

    End Sub
    Sub option37()          'Upload data to tables - All Columns

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0
        Dim tablename As String = InputBox("Enter table name", "", "???")
        Dim delimiter As String = InputBox("Enter delimiter", "", "|")
        Dim startline As String = InputBox("Enter line from which data starts", "", "1")
        Dim newtable As String = InputBox("Create new table?", "", tablename)
        Dim deleteexistingdata As String = InputBox("Delete existing data?", "", "N")
        totalfiles = dirs.Length

        If UCase(newtable) <> UCase(tablename) Then
            sql = "create table " & newtable & " as select * from " & tablename & " where rownum < 1"
            oracle_execute_non_query("ten", username, username, sql)
            'Thread.Sleep(2000)
        End If

        If UCase(deleteexistingdata) = "Y" Then
            sql = "delete from " & newtable
            oracle_execute_non_query("ten", username, username, sql)
        End If

        If totalfiles = 0 Then
            processmessage("")
            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")
        Else
            For Each dir In dirs
                tempcount = tempcount + 1
                If tempcount = 1 Then
                    uploadfiledata_without_trim(dir, username, "Y")
                Else
                    uploadfiledata_without_trim(dir, username, "Y")
                End If


                Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
                Dim conn As New OracleConnection(oradb)
                conn.Open()


                processmessage("Inserting data in to " & newtable)

                sql = "PKGMISTOOL3.IMPORT_TO_TABLE_FULL_COLUMNS"
                Dim cmd1 As New OracleCommand(sql, conn)
                cmd1.CommandType = CommandType.StoredProcedure
                cmd1.Parameters.Add("TABLENAME", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = UCase(newtable)
                cmd1.Parameters.Add("DELIMITER", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = delimiter
                cmd1.Parameters.Add("DBUSER", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = UCase(username)
                cmd1.Parameters.Add("DATA_STARTING_LINE_NO", OracleDbType.Int16, 10, Nothing, ParameterDirection.Input).Value = startline
                cmd1.ExecuteNonQuery()
                conn.Close()
            Next
        End If
        processmessage("")
        MsgBox("Data uploaded successfully", MsgBoxStyle.Information, "Done !!!")

    End Sub
    Sub option38()          'Upload data to tables - Partial Columns

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0
        Dim tablename As String = InputBox("Enter table name", "", "???")
        Dim delimiter As String = InputBox("Enter delimiter", "", "|")
        Dim startline As String = InputBox("Enter line from which data starts", "", "2")
        Dim newtable As String = InputBox("Create new table?", "", tablename)
        Dim deleteexistingdata As String = InputBox("Delete existing data?", "", "N")

        If UCase(newtable) <> UCase(tablename) Then
            sql = "create table " & newtable & " as select * from " & tablename & " where rownum < 1"
            oracle_execute_non_query("ten", username, username, sql)
            'Thread.Sleep(2000)
        End If

        If UCase(deleteexistingdata) = "Y" Then
            sql = "delete from " & newtable
            oracle_execute_non_query("ten", username, username, sql)
        End If

        totalfiles = dirs.Length
        If totalfiles = 0 Then
            processmessage("")
            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")
        Else
            For Each dir In dirs
                tempcount = tempcount + 1
                If tempcount = 1 Then
                    uploadfiledata_without_trim(dir, username, "Y")
                Else
                    uploadfiledata_without_trim(dir, username, "Y")
                End If

                Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
                Dim conn As New OracleConnection(oradb)
                conn.Open()

                processmessage("Inserting data in to " & tablename)

                sql = "PKGMISTOOL3.IMPORT_TO_TABLE_PART_COLUMNS"
                Dim cmd1 As New OracleCommand(sql, conn)
                cmd1.CommandType = CommandType.StoredProcedure
                cmd1.Parameters.Add("TABLENAME", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = UCase(newtable)
                cmd1.Parameters.Add("DELIMITER", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = delimiter
                cmd1.Parameters.Add("DBUSER", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = UCase(username)
                cmd1.Parameters.Add("DATA_STARTING_LINE_NO", OracleDbType.Int16, 10, Nothing, ParameterDirection.Input).Value = startline
                cmd1.ExecuteNonQuery()
                conn.Close()
            Next
        End If
        processmessage("")
        MsgBox("Data uploaded successfully", MsgBoxStyle.Information, "Done !!!")

    End Sub
    ' Dynamically insert into access table from oracle table
    Sub optionDynamic()  '

        Dim acess_db As String
        Dim access_db_path As String

        Dim accss_table As String

        Dim access_fields() As String
        Dim access_field As String

        Dim oracle_fields() As String
        Dim oracle_field As String

        Dim solid As String
        Dim loc_branch_code As String
        Dim loc_code As String
        Dim pick_descr As String
        Dim match As Integer
        Dim solname As String()

        'Opening Oralce connection
        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        'Reading Acess database details
        'access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        'acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        'accss_table = InputBox("Enter the access table name", "Enter valu", "LOCATION")

        'access_field1 = "LOCATION"
        'access_field2 = "LOCATIONNAME"

        access_db_path = readNthLine("C:\du1\2.txt", 0)
        acess_db = readNthLine("c:\du1\2.txt", 1)
        accss_table = readNthLine("c:\du1\2.txt", 2)
        access_field = readNthLine("c:\du1\2.txt", 3)

        access_fields = access_field.Split(",")

        Dim sql As String
        'sql = "SELECT  SUBSTR(LOC_BRANCH_CODE,2,4) SOLID ,LOC_BRANCH_CODE,LOC_LOCATION_CODE,PICK_DESCRIPTION  FROM MIG_LOCATION, MIG_BRM007 WHERE LOC_BRANCH_CODE = PICK_BRANCH_CD AND PICK_KEY_TYPE = 701 AND PICK_CODE_NUM =  LOC_LOCATION_CODE ORDER BY 1, 2,3"
        sql = readNthLine("c:\du1\2.txt", 4).Trim()

        oracle_field = sql.Remove(sql.IndexOf("FROM"), sql.Length() - sql.IndexOf("FROM")).Trim
        oracle_field = oracle_field.Replace("SELECT ", "")
        oracle_fields = oracle_field.Split(",")
        Dim orcale_field_count As Integer = readNthLine("c:\du1\2.txt", 5)
        'Getting Oracle connection
        Dim cmd As New OracleCommand(sql, oracle_conn)
        Dim dr As OracleDataReader = cmd.ExecuteReader()
        While dr.Read()

            match = 0
            'solid = ""
            'loc_branch_code = ""
            'loc_code = ""
            'pick_descr = ""

            'solid = dr(0).ToString()
            'loc_branch_code = dr(1).ToString()

            For i As Integer = 0 To orcale_field_count - 1
                oracle_fields(i) = ""

            Next

            For i As Integer = 0 To orcale_field_count - 1
                oracle_fields(i) = dr(i)

            Next
            'Handling solid issues  for makkiyad 3703,Cherukunnu 1704 and sreepuram 1703 , as per solid list makkiyad  701,cherukunnu 704, Sreepuram 176
            If dr(1) = "3703" Then
                solid = "00701"
            ElseIf dr(1) = "1704" Then
                solid = "00704"
            ElseIf dr(1) = "1703" Then
                solid = "00176"
            Else
                solid = dr(0)
                solid = solid.ToString().PadLeft(5, "0")
            End If


            loc_code = dr(2)
            pick_descr = dr(3).ToString()

            'dir1 retrives full path. so splitting the path to get the folder name 
            For Each dir1 In Directory.GetDirectories(access_db_path)
                solname = dir1.Split("\")
                If solname(solname.Length - 1).Substring(0, 5) = solid Then
                    match = match + 1  'To check the whether duplicate folder exist with the same solid
                End If
            Next

            If match > 1 Then
                MsgBox("Some conflict in SOLID. Cannot execute")

            ElseIf match = 1 Then

                If File.Exists(access_db_path & "\" & solname(solname.Length - 1).ToString() & "\" & acess_db) Then

                    Dim cnn As New OleDb.OleDbConnection
                    cnn = New OleDb.OleDbConnection

                    Dim strConnection As String
                    strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & access_db_path & "\" & solname(solname.Length - 1).ToString() & "\" & acess_db
                    cnn.ConnectionString = strConnection

                    Try
                        If Not cnn.State = ConnectionState.Open Then
                            cnn.Open()
                        End If
                    Catch ex As Exception
                        MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                    End Try

                    'Code to skip the records which are already inserted
                    Dim cmd1 As New OleDb.OleDbCommand
                    cmd1.CommandText = "Select " & access_fields(0) & " from " & accss_table & " where " & access_fields(0) & "=" & dr(2) & "and " & access_fields(1) & "='" & dr(3) & "'"
                    cmd1.Connection = cnn
                    Dim dr1 As OleDb.OleDbDataReader
                    dr1 = cmd1.ExecuteReader

                    If dr1.Read = False Then
                        Try
                            Dim cmd2 As New OleDb.OleDbCommand
                            cmd2.CommandText = "insert into " & accss_table & " (" & access_fields(0) & "," & access_fields(1) & ") values (" & dr(2) & ",'" & dr(3) & "')"
                            cmd2.Connection = cnn
                            cmd2.ExecuteNonQuery()  'Executing command
                        Catch ex As Exception
                            MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                            Exit Sub
                        End Try
                    End If
                    cnn.Close()

                End If

            Else

                MsgBox("No database found for SOLID " & solid)

            End If
            processmessage("Executing Query in file " & loc_branch_code & "\" & loc_code)
        End While
        dr.Close()
        oracle_conn.Close()
        MsgBox("Process over")
    End Sub

    Sub option801()  'Inserting to Access Location table

        Dim acess_db As String
        Dim access_db_path As String

        Dim accss_table As String

        Dim access_field1 As String
        Dim access_field2 As String

        Dim solid As String
        Dim loc_branch_code As String
        Dim loc_code As String
        Dim pick_descr As String
        Dim solname As String()


        'Opening Oralce connection
        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        'Reading Acess database details
        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database Name", "Enter valu", "NMGB.mdb")
        accss_table = "LOCATION"

        access_field1 = "LOCATION"
        access_field2 = "LOCATIONNAME"

        'Fetching soid
        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)

            'Establish connection to access database
            If File.Exists(access_db_path & "\" & solname(solname.Length - 1) & "\Server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & access_db_path & "\" & solname(solname.Length - 1) & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try


                Dim sql As String
                If solid = "00701" Then
                    sql = "SELECT  SUBSTR(LOC_BRANCH_CODE,2,4) SOLID ,LOC_BRANCH_CODE,LOC_LOCATION_CODE,PICK_DESCRIPTION  FROM MIG_LOCATION, MIG_BRM007 WHERE LOC_BRANCH_CODE = PICK_BRANCH_CD AND PICK_KEY_TYPE = 701 AND PICK_CODE_NUM =  LOC_LOCATION_CODE  AND LOC_BRANCH_CODE= 3703 ORDER BY 3"
                ElseIf solid = "00704" Then
                    sql = "SELECT  SUBSTR(LOC_BRANCH_CODE,2,4) SOLID ,LOC_BRANCH_CODE,LOC_LOCATION_CODE,PICK_DESCRIPTION  FROM MIG_LOCATION, MIG_BRM007 WHERE LOC_BRANCH_CODE = PICK_BRANCH_CD AND PICK_KEY_TYPE = 701 AND PICK_CODE_NUM =  LOC_LOCATION_CODE  AND LOC_BRANCH_CODE= 1704 ORDER BY 3"
                ElseIf solid = "00176" Then
                    sql = "SELECT  SUBSTR(LOC_BRANCH_CODE,2,4) SOLID ,LOC_BRANCH_CODE,LOC_LOCATION_CODE,PICK_DESCRIPTION  FROM MIG_LOCATION, MIG_BRM007 WHERE LOC_BRANCH_CODE = PICK_BRANCH_CD AND PICK_KEY_TYPE = 701 AND PICK_CODE_NUM =  LOC_LOCATION_CODE  AND LOC_BRANCH_CODE= 1703 ORDER BY 3"
                Else
                    sql = "SELECT  SUBSTR(LOC_BRANCH_CODE,2,4) SOLID ,LOC_BRANCH_CODE,LOC_LOCATION_CODE,PICK_DESCRIPTION  FROM MIG_LOCATION, MIG_BRM007 WHERE LOC_BRANCH_CODE = PICK_BRANCH_CD AND PICK_KEY_TYPE = 701 AND PICK_CODE_NUM =  LOC_LOCATION_CODE  AND TO_NUMBER(SUBSTR(LOC_BRANCH_CODE,2,4)) = TO_NUMBER ( '" & solid & "' ) ORDER BY 3"
                End If

                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()
                While dr.Read()
                    solid = ""
                    loc_branch_code = ""
                    loc_code = ""
                    pick_descr = ""

                    loc_branch_code = dr(1).ToString()
                    loc_code = dr(2)
                    pick_descr = dr(3).ToString()


                    'Code to skip the records which are already inserted and insert others
                    Dim cmd1 As New OleDb.OleDbCommand
                    cmd1.CommandText = "Select count (" & access_field1 & " ) as  aa from " & accss_table & " where " & access_field1 & "=" & dr(2) & " and " & access_field2 & "='" & dr(3) & "'"
                    cmd1.Connection = cnn
                    Dim dr1 As OleDb.OleDbDataReader
                    dr1 = cmd1.ExecuteReader

                    If dr1.Read = True Then
                        If dr1("aa") < 1 Then
                            Dim aa As Integer = dr1("aa")
                            Try
                                Dim cmd2 As New OleDb.OleDbCommand
                                cmd2.CommandText = "insert into " & accss_table & " (" & access_field1 & "," & access_field2 & ") values (" & dr(2) & ",'" & dr(3) & "')"
                                cmd2.Connection = cnn
                                cmd2.ExecuteNonQuery()  'Executing command
                            Catch ex As Exception
                                MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                                Exit Sub
                            End Try
                        End If
                    End If
                    dr1.Close()

                    processmessage("Executing Query in file " & dir1.ToString() & "\" & loc_code)
                End While
                dr.Close()
                cnn.Close()
            End If
        Next

        oracle_conn.Close()
        MsgBox("Process over")
    End Sub


    Private Sub option802()   'Inserting data to CIDMASTER from CEDGE_EXTRACT_CUSTOMERID table
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String

        Dim solname() As String
        Dim solid As String

        Dim cid As String
        Dim cname As String
        Dim father As String
        Dim relation As String
        Dim address1 As String
        Dim address2 As String
        Dim address3 As String
        Dim address4 As String
        Dim pincode As String
        Dim dob As Date
        Dim dobstring As String

        Dim custtype As String
        Dim title As String

        Dim match As Integer = 0

        'Opening Oralce connection
        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        'Reading Acess database details
        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        accss_table = "CIDMASTER"

        'Fetching solid and open connection to access database
        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                'Reading data from oracle table
                Dim sql As String
                sql = "SELECT DISTINCT (CID),SOLID,CNAME,FATHER,RELATION,ADDRESS1,ADDRESS2,ADDRESS3,ADDRESS4,PINCODE,DOBSTRING,CUST_TYPE,TITLE FROM CEDGE_EXTRACT_CUSTOMERID WHERE SOLID = '" & solid.Trim() & "'"
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()
                While dr.Read()
                    match = 0
                    solid = ""
                    cid = ""
                    cname = ""
                    father = ""
                    relation = ""
                    address1 = ""
                    address2 = ""
                    address3 = ""
                    address4 = ""
                    pincode = ""
                    dobstring = ""

                    custtype = ""
                    title = ""

                    solid = dr(1).ToString()
                    cid = dr(0).ToString()
                    cname = dr(2).ToString()
                    father = dr(3).ToString()
                    relation = dr(4).ToString()
                    address1 = dr(5).ToString()
                    address2 = dr(6).ToString()
                    address3 = dr(7).ToString()
                    address4 = dr(8).ToString()
                    pincode = dr(9).ToString()

                    If Len(pincode) = 8 Then
                        If pincode.Substring(6, 2) = "00" Then
                            pincode = pincode.Substring(0, 6)
                        End If
                    End If
                    dobstring = dr(10).ToString().Trim
                    If dobstring = "0" Then
                        dobstring = "01-JAN-00"
                    End If



                    custtype = dr(11).ToString()
                    title = dr(12).ToString()


                    'Code to skip the records which are already inserted and insert others
                    Dim cmd1 As New OleDb.OleDbCommand
                    cmd1.CommandText = "Select count (cid) as  aa from cidmaster where cid = '" & cid & "'"
                    cmd1.Connection = cnn
                    Dim dr1 As OleDb.OleDbDataReader
                    dr1 = cmd1.ExecuteReader

                    If dr1.Read = True Then
                        If dr1("aa") < 1 Then
                            Dim aa As Integer = dr1("aa")
                            Try
                                Dim cmd2 As New OleDb.OleDbCommand
                                dob = DateTime.ParseExact(dobstring, "dd-MMM-yy", Nothing)
                                dobstring = dob.ToString("MM-dd-yyyy")

                                'cmd2.CommandText = "INSERT INTO CIDMASTER (CID,CIDNAME,FATHER,RELATION,ADDRESS1,ADDRESS2,ADDRESS3,ADDRESS4,PINCODE,DOB) VALUES ('" & cid & "','" & cname & "','" & father & "','" & relation & "','" & address1 & "','" & address2 & "','" & address3 & "','" & address4 & "','" & pincode & "',#" & dobstring & "#)"
                                cmd2.CommandText = "INSERT INTO CIDMASTER (CID,CIDNAME,FATHER,RELATION,ADDRESS1,ADDRESS2,ADDRESS3,ADDRESS4,PINCODE,DOB,CUST_TYPE,TITLE) VALUES ('" & cid & "','" & cname & "','" & father & "','" & relation & "','" & address1 & "','" & address2 & "','" & address3 & "','" & address4 & "','" & pincode & "',#" & dobstring & "#,'" & custtype & "','" & title & "')"

                                cmd2.Connection = cnn
                                cmd2.ExecuteNonQuery()  'Executing command
                            Catch ex As Exception
                                'MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                                MsgBox(cid)
                                'Exit Sub
                            End Try
                        End If
                    End If
                    dr1.Close()
                    processmessage("Executing Query in file " & dir1.ToString() & " Writing record " & cid)
                End While
                dr.Close()
                cnn.Close()
            End If

        Next
        oracle_conn.Close()
        MsgBox("Process over")

    End Sub

    Private Sub option803() 'Inserting data to Pickup table from oralce "Pickup_tobranch" table
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String

        Dim solname() As String
        Dim solid As String

        Dim codetype As String
        Dim code As String
        Dim description As String
        Dim linkdata As String


        'Opening Oralce connection
        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        'Reading Acess database details
        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "PICKUP"

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)

            If File.Exists(dir1.ToString() & "\Server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\Server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = "SELECT CODETYPE,CODE,DESCRIPTION,LINKDATA FROM PICKUP_TOBRANCH ORDER BY CODETYPE,CODE"

                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()
                While dr.Read()

                    codetype = ""
                    code = ""
                    description = ""
                    linkdata = ""

                    code = dr(1).ToString()
                    codetype = dr(0).ToString()
                    description = dr(2).ToString()
                    linkdata = dr(3).ToString()

                    'Code to skip the records which are already inserted and insert others
                    Dim cmd1 As New OleDb.OleDbCommand
                    cmd1.CommandText = "Select count (subslno) as  aa from pickup where slno = " & codetype & "and subslno='" & code & "'"
                    cmd1.Connection = cnn
                    Dim dr1 As OleDb.OleDbDataReader
                    dr1 = cmd1.ExecuteReader

                    If dr1.Read = True Then
                        If dr1("aa") < 1 Then
                            Try
                                Dim cmd2 As New OleDb.OleDbCommand
                                cmd2.CommandText = "INSERT INTO PICKUP (SLNO,SUBSLNO,DESCRIPTION,LINKDATA) VALUES (" & codetype & ",'" & code & "','" & description & "','" & linkdata & "')"
                                cmd2.Connection = cnn
                                cmd2.ExecuteNonQuery()  'Executing command
                            Catch ex As Exception
                                MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                                Exit Sub
                            End Try
                        End If
                    End If

                    dr1.Close()
                    processmessage("Executing Query in file " & dir1.ToString() & " Writing record " & codetype & "  " & code)
                End While

                dr.Close()
                cnn.Close()

            End If

        Next
        oracle_conn.Close()
        MsgBox("Process over")
    End Sub

    Private Sub option804()  'Inserting data to Religion by taking from corresponding sols CIDMASTER
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String

        Dim cmd1 As New OleDb.OleDbCommand
        Dim cmd2 As New OleDb.OleDbCommand
        Dim cmd3 As New OleDb.OleDbCommand

        Dim dr As OleDb.OleDbDataReader
        Dim dr1 As OleDb.OleDbDataReader
        Dim dr3 As OleDb.OleDbDataReader

        Dim recordslno As Integer
        Dim totalrecords As Integer
        Dim noofrecords As Integer
        Dim custid As String

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        accss_table = "RELIGION"


        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                cmd3.Connection = cnn
                cmd3.CommandText = "SELECT IIF(MAX(SLNO) IS NULL ,0,MAX(SLNO)) FROM RELIGION"
                dr3 = cmd3.ExecuteReader
                If dr3.Read = True Then
                    recordslno = dr3(0)
                Else
                    recordslno = 0
                End If
                dr3.Close()

                cmd1.Connection = cnn
                'cmd1.CommandText = "SELECT DISTINCT CID FROM CIDMASTER"
                cmd1.CommandText = "SELECT DISTINCT CID FROM CIDMASTER WHERE CUST_TYPE IN ('010101','010102','010103','010104','010105','010106','010107','010110','010201','010202','01','010109','010108')"
                dr = cmd1.ExecuteReader
                While dr.Read()
                    totalrecords = totalrecords + 1
                    custid = dr("CID").ToString
                    cmd2.Connection = cnn
                    cmd2.CommandText = "SELECT COUNT(1) FROM RELIGION WHERE CUSTOMERID = '" & custid & "'"
                    dr1 = cmd2.ExecuteReader
                    If dr1.Read = True Then
                        tempcount = dr1(0)
                    End If
                    dr1.Close()

                    If tempcount = 0 Then
                        recordslno = recordslno + 1
                        noofrecords = noofrecords + 1
                        cmd3.CommandText = "INSERT INTO RELIGION (SLNO,CUSTOMERID) VALUES (" & recordslno & ",'" & custid & "')"
                        cmd3.Connection = cnn
                        cmd3.ExecuteNonQuery()

                    End If
                    processmessage("Writing in file " & solid & " Record NO : " & custid)
                End While
                dr.Close()
                cnn.Close()
            End If
        Next
        MsgBox("Religion code inserted successfully")
    End Sub

    Private Sub option805()   'Updating religion by taking data from banc724 backups
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim solid_int As Integer

        Dim custid As String
        Dim religion As String
        Dim caste As String
        Dim custid_10 As String

        Dim cmd1 As New OleDb.OleDbCommand


        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        accss_table = "RELIGION"

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            Dim totalrecords As Integer = 0
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim cmd As New OleDb.OleDbCommand
                Dim dr As OleDb.OleDbDataReader
                cmd.CommandText = "SELECT CUSTOMERID FROM RELIGION WHERE LEFT(CUSTOMERID,1) = 6"
                cmd.Connection = cnn
                dr = cmd.ExecuteReader()

                While dr.Read()
                    custid_10 = ""
                    custid = ""
                    custid = dr(0)
                    custid_10 = custid.Substring(0, 10)

                    sql = " SELECT  DECODE (CUST_RELIGION,'3','1', '4','2', '5','3'), DECODE( CUST_CAST,'4','3','5','4', '6','2'),CBS_ID  FROM MIG_BRM001 WHERE BRANCHID = " & solid_int & "  AND CBS_ID=' " & custid_10 & " '"
                    Dim cmd2 As New OracleCommand(sql, oracle_conn)
                    Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
                    If dr2.Read() Then
                        religion = ""
                        caste = ""
                        religion = dr2(0).ToString()
                        caste = dr2(1).ToString()


                        'Updating data to access table
                        Try

                            'Executing command
                            If religion <> "" Then
                                Dim cmd3 As New OleDb.OleDbCommand
                                cmd3.CommandText = "UPDATE  RELIGION  SET RELIGIONCODE = '" & religion & "' WHERE CUSTOMERID= '" & custid & "' "
                                cmd3.Connection = cnn
                                cmd3.ExecuteNonQuery()
                            End If
                            If caste <> "" Then
                                Dim cmd3 As New OleDb.OleDbCommand
                                cmd3.CommandText = "UPDATE  RELIGION  SET CASTCODE = '" & caste & "' WHERE CUSTOMERID = '" & custid & "' "
                                cmd3.Connection = cnn
                                cmd3.ExecuteNonQuery()  'Executing command
                            End If

                        Catch ex As Exception
                            MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                            Exit Sub
                        End Try

                        processmessage("Writing in file" & solid & "Record " & custid)
                    End If
                    dr2.Close()
                End While
                dr.Close()
                cnn.Close()
            End If
        Next
        oracle_conn.Close()
        MsgBox("Process over")
    End Sub


    Private Sub option806()   'Inserting into Branch master table
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String

        Dim dtname As String
        Dim roname As String
        Dim sol As String
        Dim version As Integer = 1


        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        accss_table = "BRMASTER"

        'Fetching solid and open access database connection
        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)


            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try
                If solname(solname.Length - 1).Contains("_") Then
                    solname = solname(solname.Length - 1).Split("_")
                    sol = solname(0)
                    dtname = solname(1)
                    roname = solname(2)
                    solid = solname(3)

                    Dim cmd3 As New OleDb.OleDbCommand
                    cmd3.CommandText = "INSERT INTO BRMASTER (SOLID,SOLNAME,DISTRICT,RO,VERSION) VALUES (val('" & sol & " ') , '" & solid & "' , '" & dtname & "','" & roname & "'," & version & ")"
                    cmd3.Connection = cnn
                    cmd3.ExecuteNonQuery()
                    cnn.Close()
                End If

            End If

        Next

        MsgBox("Branch master data inserted")
    End Sub
    Private Sub option807()   'Inserting to ACMASTER from deposit shadow file
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim shadowfile_path As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String

        Dim tempcount As Integer
        Dim tempvar As String
        Dim acno As String
        Dim custid As String
        Dim custname As String
        Dim productcode As String
        Dim closedate As String
        Dim amount As String
        Dim recordcount As Integer
        Dim opendatestring As String
        Dim opendate As Date

        Dim dr As OleDb.OleDbDataReader
        Dim cmd As New OleDb.OleDbCommand

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du1")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        shadowfile_path = InputBox("Enter shadow file name (with full path)", "Enter valu", "c:\du")
        accss_table = "ACMASTER"

        'Fetching solid and open access database connection
        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)

            If File.Exists(shadowfile_path & "\" & sol & ".txt") Then

                If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                    Dim cnn As New OleDb.OleDbConnection
                    cnn = New OleDb.OleDbConnection

                    Dim strConnection As String
                    strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                    cnn.ConnectionString = strConnection

                    Try
                        If Not cnn.State = ConnectionState.Open Then
                            cnn.Open()
                        End If
                    Catch ex As Exception
                        MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                    End Try

                    Dim sr As StreamReader = New StreamReader(shadowfile_path & "\" & sol & ".txt")
                    tempcount = 0
                    Do While sr.Peek() >= 0
                        tempcount = tempcount + 1
                        tempvar = sr.ReadLine()
                        acno = tempvar.Substring(3, 17)
                        acno = acno.TrimStart("0")
                        custid = tempvar.Substring(25, 17)
                        custid = custid.TrimStart("0")
                        custname = tempvar.Substring(42, 60)
                        custname = custname.Trim
                        custname = custname.TrimEnd(".")
                        custname = custname.Trim

                        custname = custname.Replace("'", "")
                        custname = custname.Replace(")", "")
                        custname = custname.Replace("(", "")

                        productcode = tempvar.Substring(224, 8)
                        productcode = productcode.Trim
                        closedate = tempvar.Substring(253, 8)
                        closedate = closedate.Trim
                        amount = tempvar.Substring(153, 17)
                        amount = amount.TrimStart("0")

                        opendatestring = tempvar.Substring(102, 8)
                        opendate = Date.Parse(String.Concat(opendatestring.Substring(0, 2), "-", opendatestring.Substring(2, 2), "-", opendatestring.Substring(4, 4)))

                        If amount = "" Then
                            amount = "0"
                        End If
                        amount = (Long.Parse(amount) / 1000).ToString
                        tempvar = tempvar.Replace("'", "`")
                        If closedate = "" Then
                            sql = "select count(1) from acmaster where acno = '" & acno & "'"
                            cmd.Connection = cnn
                            cmd.CommandText = sql
                            dr = cmd.ExecuteReader()
                            If dr.Read = True Then
                                recordcount = dr(0)
                            End If
                            dr.Close()
                            If recordcount = 0 Then
                                sql = "insert into acmaster (acno,custid,custname,productcode,amount,opendate) values ('" & acno & "','" & custid & "','" & custname & "','" & productcode & "'," & amount & ",#" & opendate & "# )"
                                Try
                                    cmd.CommandText = sql
                                    cmd.Connection = cnn
                                    cmd.ExecuteNonQuery()
                                Catch ex As Exception
                                    MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                                End Try
                            End If
                        End If

                        processmessage("Retrieving data: " & sol & "Account No - " & acno)
                    Loop
                    sr.Close()
                    cnn.Close()
                End If
            End If

        Next

        MsgBox("Deposit shadow file data updated successfully")
    End Sub
    Private Sub option808()   'Inserting to ACMASTER from loan shadow file
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim shadowfile_path As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String

        Dim tempcount As Integer
        Dim tempvar As String
        Dim acno As String
        Dim custid As String
        Dim custname As String
        Dim productcode As String
        Dim closedate As String
        Dim recordcount As Integer
        Dim cmd As New OleDb.OleDbCommand
        Dim dr As OleDb.OleDbDataReader
        Dim amount As String

        Dim opendatestring As String
        Dim opendate As Date

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du1")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        shadowfile_path = InputBox("Enter Loan shadow file name (with full path)", "Enter valu", "c:\du")
        accss_table = "ACMASTER"

        'Fetching solid and open access database connection
        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")

            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)

            'Check shadow file exists or not
            If File.Exists(shadowfile_path & "\" & sol & ".txt") Then

                If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                    Dim cnn As New OleDb.OleDbConnection
                    cnn = New OleDb.OleDbConnection

                    Dim strConnection As String
                    strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                    cnn.ConnectionString = strConnection

                    Try
                        If Not cnn.State = ConnectionState.Open Then
                            cnn.Open()
                        End If
                    Catch ex As Exception
                        MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                    End Try

                    Dim sr As StreamReader = New StreamReader(shadowfile_path & "\" & sol & ".txt")
                    tempcount = 0
                    Do While sr.Peek() >= 0
                        tempcount = tempcount + 1
                        tempvar = sr.ReadLine()
                        acno = tempvar.Substring(3, 17)
                        acno = acno.TrimStart("0")

                        custid = tempvar.Substring(25, 17)
                        custid = custid.TrimStart("0")

                        custname = tempvar.Substring(42, 60)
                        custname = custname.Trim
                        custname = custname.TrimEnd(".")
                        custname = custname.Trim
                        custname = custname.Replace("'", "")
                        custname = custname.Replace(")", "")
                        custname = custname.Replace("(", "")

                        productcode = tempvar.Substring(196, 8)
                        productcode = productcode.Trim

                        closedate = tempvar.Substring(225, 8)
                        closedate = closedate.Trim

                        amount = tempvar.Substring(270, 17)
                        amount = amount.TrimStart("0")

                        opendatestring = tempvar.Substring(102, 8)
                        If opendatestring = "99999999" Then
                            opendatestring = "01011976"
                        End If
                        opendate = Date.Parse(String.Concat(opendatestring.Substring(0, 2), "-", opendatestring.Substring(2, 2), "-", opendatestring.Substring(4, 4)))

                        If amount = "" Then
                            amount = "0"
                        End If
                        amount = (Long.Parse(amount) / 1000).ToString

                        tempvar = tempvar.Replace("'", "`")

                        If closedate = "" Then
                            sql = "select count(1) from acmaster where acno = '" & acno & "'"
                            cmd.Connection = cnn
                            cmd.CommandText = sql
                            dr = cmd.ExecuteReader()
                            If dr.Read = True Then
                                recordcount = dr(0)
                            End If
                            dr.Close()
                            If recordcount = 0 Then
                                sql = "insert into acmaster (acno,custid,custname,productcode,amount,opendate) values ('" & acno & "','" & custid & "','" & custname & "','" & productcode & "'," & amount & ",#" & opendate & "#)"
                                Try
                                    cmd.CommandText = sql
                                    cmd.Connection = cnn
                                    cmd.ExecuteNonQuery()
                                Catch ex As Exception
                                    MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                                End Try
                            End If
                        End If
                        processmessage("Retrieving data: " & sol & "Account  No - " & acno)

                    Loop
                    sr.Close()
                    cnn.Close()
                End If
            End If

        Next

        MsgBox("Loan shadow file data updated successfully")
    End Sub

    Private Sub option809()   'Inserting data for NRE account
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String

        Dim recordslno As Integer = 0
        Dim nextrecordslno As Integer = 0
        Dim noofrecords As Integer = 0

        Dim custid As String

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        accss_table = "NRECODE"

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim cmd2 As New OleDb.OleDbCommand
                Dim cmd3 As New OleDb.OleDbCommand
                Dim cmd4 As New OleDb.OleDbCommand

                Dim dr1 As OleDb.OleDbDataReader
                Dim dr2 As OleDb.OleDbDataReader
                Dim dr3 As OleDb.OleDbDataReader

                cmd3.Connection = cnn
                cmd3.CommandText = "SELECT IIF(MAX(SLNO) IS NULL ,0,MAX(SLNO)) FROM NRECODE"
                dr3 = cmd3.ExecuteReader
                If dr3.Read = True Then
                    recordslno = dr3(0)
                Else
                    recordslno = 0
                End If
                dr3.Close()

                cmd2.Connection = cnn
                cmd2.CommandText = "SELECT DISTINCT CUSTID FROM ACMASTER WHERE PRODUCTCODE IN (SELECT SUBSLNO FROM PICKUP WHERE SLNO IN (9,10))"
                dr1 = cmd2.ExecuteReader

                While dr1.Read()
                    tempcount = 0
                    custid = dr1("CUSTID")
                    nextrecordslno = nextrecordslno + 1
                    cmd3.Connection = cnn
                    cmd3.CommandText = "SELECT COUNT(1) FROM NRECODE WHERE CID = '" & custid & "'"
                    dr2 = cmd3.ExecuteReader

                    If dr2.Read = True Then
                        tempcount = dr2(0)
                    End If
                    dr2.Close()

                    If tempcount = 0 Then
                        recordslno = recordslno + 1
                        noofrecords = noofrecords + 1
                        cmd4.CommandText = "INSERT INTO NRECODE (SLNO,CID) VALUES (" & recordslno & ",'" & custid & "')"
                        cmd4.Connection = cnn
                        cmd4.ExecuteNonQuery()
                    End If
                    processmessage("Inserting for : " & sol & " Writing Customer ID: " & custid)
                End While

                dr1.Close()
                cnn.Close()
            End If

        Next

        MsgBox("NRE data inserted successfully")
    End Sub

    Private Sub option810()   'Inserting data for Staff account
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim custid As String

        Dim recordslno As Integer = 0
        Dim nextrecordslno As Integer = 0
        Dim noofrecords As Integer = 0

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "STAFFCODE"

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim cmd1 As New OleDb.OleDbCommand
                Dim cmd2 As New OleDb.OleDbCommand
                Dim cmd3 As New OleDb.OleDbCommand
                Dim cmd4 As New OleDb.OleDbCommand

                Dim dr1 As OleDb.OleDbDataReader
                Dim dr2 As OleDb.OleDbDataReader
                Dim dr3 As OleDb.OleDbDataReader

                cmd1.Connection = cnn
                cmd1.CommandText = "SELECT IIF(MAX(SLNO) IS NULL ,0,MAX(SLNO)) FROM STAFFCODE"
                dr3 = cmd1.ExecuteReader
                If dr3.Read = True Then
                    recordslno = dr3(0)
                Else
                    recordslno = 0
                End If
                dr3.Close()


                cmd2.Connection = cnn
                cmd2.CommandText = "SELECT DISTINCT CUSTID FROM ACMASTER WHERE PRODUCTCODE IN (SELECT SUBSLNO FROM PICKUP WHERE SLNO = 15)"
                dr1 = cmd2.ExecuteReader
                While dr1.Read()
                    custid = dr1("CUSTID")
                    nextrecordslno = nextrecordslno + 1

                    cmd3.Connection = cnn
                    cmd3.CommandText = "SELECT COUNT(1) FROM STAFFCODE WHERE CID = '" & custid & "'"
                    dr2 = cmd3.ExecuteReader
                    If dr2.Read = True Then
                        tempcount = dr2(0)
                    End If
                    dr2.Close()

                    If tempcount = 0 Then
                        recordslno = recordslno + 1
                        noofrecords = noofrecords + 1
                        cmd4.CommandText = "INSERT INTO STAFFCODE (SLNO,CID) VALUES (" & recordslno & ",'" & custid & "')"
                        cmd4.Connection = cnn
                        cmd4.ExecuteNonQuery()
                        'TextBox9.Text = recordslno & " new records found"
                    End If
                    processmessage("Inserting for : " & sol & " Writing Customer ID : " & custid)

                End While
                dr1.Close()
                cnn.Close()

            End If

        Next

        MsgBox("Staff Account numbers inserted successfully")
    End Sub

    Private Sub option811()   'Inserting data for Customer category
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim solid_int As Integer

        Dim custname As String
        Dim categorytype As String
        Dim categorygroup As String
        Dim cbscustid As String
        Dim custid_11 As String


        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "CUSTCATEGORY"

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = " SELECT  CBS_ID,CUST_FIRST_NAME,DECODE (CUST_TYPE, '2','BP','4','HP'), DECODE (CUST_SPEC_ATTR_CODE, '9' ,'GL','8','JA','10','EX','7','SW')  FROM MIG_BRM001  WHERE  BRANCHID = " & solid_int & "AND (CUST_TYPE = '2' OR CUST_TYPE = '4' OR CUST_SPEC_ATTR_CODE BETWEEN 7 AND 10) AND CBS_ID IS NOT NULL "

                'Retriving data from Oracle
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr2 As OracleDataReader = cmd.ExecuteReader()

                While dr2.Read()
                    cbscustid = ""
                    custname = ""
                    categorytype = ""
                    categorygroup = ""

                    cbscustid = dr2(0).ToString()
                    custname = dr2(1).ToString()
                    categorytype = dr2(2).ToString()
                    categorygroup = dr2(3).ToString()

                    Dim sql1 As String
                    sql1 = "SELECT CID FROM CEDGE_EXTRACT_CUSTOMERID WHERE CID_10='" & cbscustid & "'"
                    Dim cmd5 As New OracleCommand(sql1, oracle_conn)
                    Dim dr5 As OracleDataReader = cmd5.ExecuteReader()
                    If dr5.Read() Then
                        custid_11 = ""
                        custid_11 = dr5(0)

                        Dim cmd3 As New OleDb.OleDbCommand
                        Dim dr3 As OleDb.OleDbDataReader
                        cmd3.Connection = cnn
                        cmd3.CommandText = "SELECT COUNT(1) FROM CUSTCATEGORY WHERE CID = '" & custid_11 & "'"
                        dr3 = cmd3.ExecuteReader

                        If dr3.Read = True Then
                            tempcount = dr3(0)
                        End If
                        dr3.Close()
                        If tempcount = 0 Then
                            Dim cmd4 As New OleDb.OleDbCommand
                            cmd4.CommandText = "INSERT INTO CUSTCATEGORY (CID,CIDNAME,CATEGORYTYPE,CATEGORYGROUP) VALUES ( '" & custid_11 & "','" & custname & "','" & categorytype & "','" & categorygroup & "')"
                            cmd4.Connection = cnn
                            cmd4.ExecuteNonQuery()
                        End If
                    End If
                    dr5.Close()
                    processmessage("Inserting data for : " & sol & " Writing Customer ID: " & cbscustid)
                End While
                dr2.Close()
            End If
            cnn.Close()

        Next

        oracle_conn.Close()
        MsgBox("Customer category updated from Banc724 successfully")
    End Sub
    Private Sub option812()   'Inserting data to Citycode1 from CIDmaster
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solname() As String
        Dim solid As String


        Dim address3 As String = ""
        Dim pincode As String = ""

        ''Opening Oralce connection
        'Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        'Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        'oracle_conn.Open()

        'Reading Acess database details
        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        accss_table = "Citycode1"

        'Fetching solid and open connection to access database
        For Each dir1 In Directory.GetDirectories(access_db_path)
            Dim recordslno As Integer = 0
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim cmd1 As New OleDb.OleDbCommand
                Dim cmd2 As New OleDb.OleDbCommand
                Dim cmd3 As New OleDb.OleDbCommand

                Dim dr As OleDb.OleDbDataReader
                Dim dr1 As OleDb.OleDbDataReader
                Dim dr3 As OleDb.OleDbDataReader

                cmd3.Connection = cnn
                cmd3.CommandText = "SELECT IIF(MAX(SLNO) IS NULL ,0,MAX(SLNO)) FROM CITYCODEI"
                dr3 = cmd3.ExecuteReader
                If dr3.Read = True Then
                    recordslno = dr3(0)
                Else
                    recordslno = 0
                End If
                dr3.Close()



                cmd1.Connection = cnn
                cmd1.CommandText = "SELECT DISTINCT ADDRESS2 & ' : ' & ADDRESS3 AS ADDRESS3,PINCODE FROM CIDMASTER WHERE PINCODE IN (SELECT LINKDATA FROM PICKUP WHERE SLNO = 7)"
                dr = cmd1.ExecuteReader
                While dr.Read()
                    address3 = ""
                    pincode = ""

                    address3 = dr("address3").ToString
                    pincode = dr("pincode")

                    cmd2.Connection = cnn
                    cmd2.CommandText = "SELECT COUNT(1) FROM CITYCODEI WHERE ADDRESS3 = '" & address3 & "' AND PINCODE = '" & pincode & "'"
                    dr1 = cmd2.ExecuteReader

                    If dr1.Read = True Then
                        tempcount = dr1(0)
                    End If
                    dr1.Close()

                    If tempcount = 0 Then
                        recordslno = recordslno + 1
                        Try
                            cmd3.CommandText = "INSERT INTO CITYCODEI (SLNO,ADDRESS3,PINCODE) VALUES (" & recordslno & ",'" & address3 & "','" & pincode & "')"
                            cmd3.Connection = cnn
                            cmd3.ExecuteNonQuery()
                        Catch
                            MsgBox(address3)
                        End Try

                    End If
                    processmessage("Executing Query in file " & solid & ", Writing : " & address3)
                End While
                dr.Close()
                cnn.Close()
            End If
        Next
        'oracle_conn.Close()
        MsgBox("Inserted into Citycode 1. Process Completed")

    End Sub

    Private Sub option813()   'Inserting data to Citycode2 from CIDmaster
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solname() As String
        Dim solid As String

        'Reading Acess database details
        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        accss_table = "Citycode2"

        'Fetching solid and open connection to access database
        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim cmd3 As New OleDb.OleDbCommand
                Dim dr3 As OleDb.OleDbDataReader
                Dim recordslno As Integer

                Dim custid As String
                cmd3.Connection = cnn
                cmd3.CommandText = "SELECT IIF(MAX(SLNO) IS NULL ,0,MAX(SLNO)) FROM CITYCODEII"
                dr3 = cmd3.ExecuteReader
                If dr3.Read = True Then
                    recordslno = dr3(0)
                Else
                    recordslno = 0
                End If
                dr3.Close()

                Dim cmd1 As New OleDb.OleDbCommand
                Dim dr As OleDb.OleDbDataReader

                cmd1.Connection = cnn
                cmd1.CommandText = "SELECT CID FROM CIDMASTER WHERE PINCODE NOT IN (SELECT LINKDATA FROM PICKUP WHERE SLNO = 7)"
                dr = cmd1.ExecuteReader
                While dr.Read()
                    custid = dr("cid")

                    Dim cmd2 As New OleDb.OleDbCommand
                    Dim dr1 As OleDb.OleDbDataReader

                    cmd2.Connection = cnn
                    cmd2.CommandText = "SELECT COUNT(1) FROM CITYCODEII WHERE CID = '" & custid & "'"
                    dr1 = cmd2.ExecuteReader
                    If dr1.Read = True Then
                        tempcount = dr1(0)
                    End If
                    dr1.Close()

                    If tempcount = 0 Then
                        recordslno = recordslno + 1
                        cmd3.CommandText = "INSERT INTO CITYCODEII (SLNO,CID) VALUES (" & recordslno & ",'" & custid & "')"
                        cmd3.Connection = cnn
                        cmd3.ExecuteNonQuery()
                    End If
                    processmessage("Executing Query in file " & solid & ", Writing : " & custid)
                End While
                dr.Close()
                cnn.Close()

            End If
        Next
        MsgBox("City code 2 inserted successfully. Process over")
    End Sub

    Private Sub option814()   'Inserting data to Guardianmaster from CIDmaster
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solname() As String
        Dim solid As String
        Dim dob As Date

        Dim cid As String

        'Reading Acess database details
        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database name", "Enter valu", "NMGB.mdb")
        accss_table = "Guardianmaster"

        'Fetching solid and open connection to access database
        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim cmd3 As New OleDb.OleDbCommand
                Dim dr3 As OleDb.OleDbDataReader
                Dim recordslno As Integer = 0

                cmd3.Connection = cnn
                cmd3.CommandText = "SELECT IIF(MAX(SLNO) IS NULL ,0,MAX(SLNO)) FROM GUARDIANMASTER"
                dr3 = cmd3.ExecuteReader
                If dr3.Read = True Then
                    recordslno = dr3(0)
                Else
                    recordslno = 0
                End If
                dr3.Close()

                Dim cmd1 As New OleDb.OleDbCommand
                Dim dr As OleDb.OleDbDataReader

                cmd1.Connection = cnn
                cmd1.CommandText = "SELECT CID,DOB FROM CIDMASTER WHERE DateDiff ('m', dob , date() )  between 1 and 216 AND CUST_TYPE IN ('010101','010102','010103','010104','010105','010106','010107','010110','010201','010202','01','010109','010108') "
                dr = cmd1.ExecuteReader
                While dr.Read()
                    cid = ""
                    dob = "01/01/1900"
                    cid = dr(0)
                    dob = dr(1)
                    Dim cmd10 As New OleDb.OleDbCommand
                    Dim dr10 As OleDb.OleDbDataReader

                    cmd10.Connection = cnn
                    cmd10.CommandText = "SELECT MIN (OPENDATE) FROM ACMASTER WHERE CUSTID='" & cid & "'"
                    dr10 = cmd10.ExecuteReader
                    Dim opendate As Date
                    Dim opendatestring As String = ""
                    While dr10.Read()
                        opendate = "01/01/1900"
                        opendatestring = dr10(0).ToString()
                        If opendatestring <> "" Then
                            opendate = Date.Parse(opendatestring)
                        End If
                    End While
                    dr10.Close()

                    If opendatestring = "" Then
                        opendate = dob
                    End If

                    If opendate > dob Then
                        Dim cmd2 As New OleDb.OleDbCommand
                        Dim dr1 As OleDb.OleDbDataReader

                        cmd2.Connection = cnn
                        cmd2.CommandText = "SELECT COUNT(1) FROM GUARDIANMASTER WHERE CID = '" & cid & "'"
                        dr1 = cmd2.ExecuteReader
                        If dr1.Read = True Then
                            tempcount = dr1(0)
                        End If
                        dr1.Close()

                        If tempcount = 0 Then
                            recordslno = recordslno + 1
                            cmd3.CommandText = "INSERT INTO GUARDIANMASTER (SLNO,CID) VALUES (" & recordslno & ",'" & cid & "')"
                            cmd3.Connection = cnn
                            cmd3.ExecuteNonQuery()
                        End If
                    End If
                    processmessage("Writing in file " & solid & " Record NO : " & cid)

                End While
                dr.Close()
                cnn.Close()

            End If
        Next

    End Sub
    Private Sub option815()

        If checkaccountfile("C:/du", Today) = 1 Then
            MsgBox("File extracted successfully for date " & Today)
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-1)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-1))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-2)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-2))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-3)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-3))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-4)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-4))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-5)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-5))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-6)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-6))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-7)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-7))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-8)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-8))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-9)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-9))
            Exit Sub
        ElseIf checkaccountfile("C:/du", Today.AddDays(-10)) = 1 Then
            MsgBox("File extracted successfully for date " & Today.AddDays(-10))
            Exit Sub
        Else
            MsgBox("No DEP_Shadow_file.txt.gz found for the last 10 days")
            Exit Sub
        End If


    End Sub
    Private Sub option816()
        Dim path_folder As String
        Dim databasefolder As String
        Dim ipaddress As String

        path_folder = InputBox("Enter the path in which solid created", "Enter value", "c:\du")
        databasefolder = InputBox("Enter databasename with full path", "Enter value", "c:\du")

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each foldername As String In Directory.GetDirectories(path_folder)
            Dim solid() As String = Path.GetFullPath(foldername).Split("\")
            Dim solname As String = solid(solid.Length - 1)
            solname = solname.ToString.Substring(2, 3)

            Dim sql As String
            sql = "select ip from sol_ip where solid ='" & solname & "'"

            'Retriving data from Oracle
            Dim cmd As New OracleCommand(sql, oracle_conn)
            Dim dr As OracleDataReader = cmd.ExecuteReader()

            While dr.Read()
                ipaddress = ""
                ipaddress = dr(0)


                Dim FILE_NAME As String = foldername & "\client\param.txt"
                If System.IO.File.Exists(FILE_NAME) = False Then
                    System.IO.File.Create(FILE_NAME).Dispose()
                Else
                    File.Delete(FILE_NAME)
                End If
                Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)
                objWriter.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\DLL\")
                objWriter.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\DB")
                objWriter.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\BR\")
                objWriter.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\")
                objWriter.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\UR\")

                objWriter.WriteLine("C:\NMGB\Reports\")
                objWriter.WriteLine("C:\NMGB\")
                objWriter.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\BD")
                objWriter.WriteLine("C:\NMGB\Backup\")

                objWriter.WriteLine("Line 1 - Server DLL Path			[Getserverdllpath]")
                objWriter.WriteLine("Line 2 - Server Database Path			[Getserverdatabasepath]")
                objWriter.WriteLine("Line 3 - Server Bank Report Path		[Getserverbankreportpath]")
                objWriter.WriteLine("Line 4 - Server CEDGE Report Path		[Getservercedgereportpath]")
                objWriter.WriteLine("Line 5 - Server report upload path		[GetServerUploadPath]")
                objWriter.WriteLine("Line 6 - Node Report Path			[GetReportPath]")
                objWriter.WriteLine("Line 7 - Node root path				[GetNodeRootPath]")

                objWriter.WriteLine("nmgb_mig_002.GetConnectionString()")
                objWriter.WriteLine("nmgb_mig_002.Getserverdllpath()")
                objWriter.WriteLine("nmgb_mig_002.Getserverdatabasepath()")
                objWriter.WriteLine("nmgb_mig_002.Getserverbankreportpath()")
                objWriter.WriteLine("nmgb_mig_002.Getservercedgereportpath()")
                objWriter.WriteLine("nmgb_mig_002.GetServerUploadPath()")
                objWriter.WriteLine("nmgb_mig_002.GetReportPath()")
                objWriter.WriteLine("nmgb_mig_002.GetNodeRootPath()")
                objWriter.Close()

                'Writing file for displaying IP address in Server folder
                Dim FILE_NAME1 As String = foldername & "\server\ipaddress.txt"
                If System.IO.File.Exists(FILE_NAME1) = False Then
                    System.IO.File.Create(FILE_NAME1).Dispose()
                Else
                    File.Delete(FILE_NAME1)
                End If

                Dim objWriter1 As New System.IO.StreamWriter(FILE_NAME1, True)
                objWriter1.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\BR")
                objWriter1.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\BD")
                objWriter1.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\DB")
                objWriter1.WriteLine("\\" & ipaddress & "\ftp\cdc\reports")
                objWriter1.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\DLL")
                objWriter1.WriteLine("\\" & ipaddress & "\ftp\cdc\reports\dit\MT\UR")
                objWriter1.Close()

            End While
            dr.Close()
            If File.Exists(foldername & "\server\nmgb.mdb") Then
                File.Delete(foldername & "\server\nmgb.mdb")
            End If
            'File.Create(foldername & "\server\nmgb.mdb")
            File.Copy(databasefolder, foldername & "\server\nmgb.mdb")
            processmessage("writing in folder" & foldername)

        Next
        MsgBox("Process over. Param file, IP address file and database file inserted")
    End Sub
    Private Sub option817()

        Dim sourcepath As String
        Dim destinationpath As String
        Dim subfolderflag As String
        Dim contentsonlyflag As String

        sourcepath = InputBox("Enter the source folder with path", "Enter value", "c:\du1")
        If Directory.Exists(sourcepath) = False Then
            MsgBox("No folder found in " & sourcepath & " sourcepath.")
            Exit Sub
        End If
        destinationpath = InputBox("Enter the destination folder with path", "Enter value", "c:\du")
        If Directory.Exists(sourcepath) = False Then
            MsgBox("No folder found in " & sourcepath & " sourcepath.")
            Exit Sub
        End If
        subfolderflag = InputBox("Include subfolder(Y/N/B)", "Enter value", "Y")
        If subfolderflag.ToUpper <> "Y" And subfolderflag.ToUpper <> "N" And subfolderflag.ToUpper <> "B" Then
            MsgBox("Enter valid option(Y/N/B)")
        End If

        contentsonlyflag = InputBox("Contents only(Y/N/)", "Enter value", "Y")
        If contentsonlyflag.ToUpper <> "Y" And contentsonlyflag.ToUpper <> "N" Then
            MsgBox("Enter valid option(Y/N)")
        End If

        If contentsonlyflag.ToUpper() = "Y" And subfolderflag = "Y" Then
            For Each dir1 As String In Directory.GetDirectories(destinationpath)

                copyfolder(sourcepath, Path.GetFullPath(dir1), "Y", "Y")

                For Each file1 As String In Directory.GetFiles(sourcepath)

                    FileCopy(Path.GetFullPath(file1), destinationpath)
                Next
                processmessage("writing in file " & dir1)
            Next


        End If
        MsgBox("Process Over. Client and Server folders copied for creating Setup")
    End Sub
    Private Sub copyfolder(ByVal source, ByVal destination, ByVal includesubfolder, ByVal contentsonlyflag)

        If contentsonlyflag.ToString.ToUpper = "N" Then
            Dim foldername() As String
            foldername = source.ToString.Split("\")
            System.IO.Directory.CreateDirectory(destination & "\" & foldername(foldername.Length() - 1))
            destination = destination & "\" & foldername(foldername.Length() - 1)
        End If

        For Each dir1 As String In Directory.GetDirectories(source)
            Dim directorypath() As String = dir1.Split("\")
            Dim directoryname As String = directorypath(directorypath.Length - 1)

            If Directory.Exists(destination & "\" & directoryname) Then
                'Directory.Delete(destination & "\" & directoryname)
            Else
                System.IO.Directory.CreateDirectory(destination & "\" & directoryname)
                copyfolder(dir1, destination & "\" & directoryname, "Y", "Y")
            End If
        Next
        For Each file1 As String In Directory.GetFiles(source)
            Dim filepath() As String = file1.Split("\")
            Dim filename As String = filepath(filepath.Length - 1)

            If File.Exists(destination & "\" & filename) Then
                File.Delete(destination & "\" & filename)
            Else
                FileCopy(Path.GetFullPath(file1), destination & "\" & filename)
            End If

        Next

    End Sub

    Private Sub option818()    'Inserting from nre  file
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim solid_int As Integer

        Dim custid As String
        Dim country As String

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "NRECODE"

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = "select solid,cid,country from nri where solid='" & solid_int & "'"
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()

                While dr.Read()
                    custid = ""
                    country = ""

                    custid = dr(1).ToString()
                    country = dr(2).ToString()

                    Dim cmd1 As New OleDb.OleDbCommand
                    Dim dr1 As OleDb.OleDbDataReader
                    cmd1.Connection = cnn
                    cmd1.CommandText = "SELECT COUNT(1) FROM PICKUP WHERE SLNO = 8  AND SUBSLNO = '" & country.ToUpper() & "'"
                    dr1 = cmd1.ExecuteReader
                    Dim temp As Integer = 0
                    If dr1.Read = True Then
                        temp = dr1(0)
                    End If
                    dr1.Close()
                    If temp = 1 Then
                        Dim cmd2 As New OleDb.OleDbCommand
                        cmd2.Connection = cnn
                        cmd2.CommandText = "UPDATE NRECODE SET COUNTRYCODE ='" & country & "' WHERE CID='" & custid & "'"
                        cmd2.ExecuteNonQuery()
                    End If

                End While
                dr.Close()
                cnn.Close()
            End If
        Next

        oracle_conn.Close()
        MsgBox("Process Completed")
    End Sub
    Private Sub option819() 'Inserting from deceased file
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim solid_int As Integer

        Dim custid As String
        Dim dateofdeathstr As String
        Dim dateofdeath As Date
        Dim noofrecords As Integer

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "DECEASECODE"

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = "select solid,custid,deceased_date from deceased where solid='" & solid_int & "'"
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()

                While dr.Read()
                    custid = ""
                    dateofdeathstr = ""

                    custid = dr(1).ToString()
                    dateofdeathstr = dr(2).ToString()
                    dateofdeath = DateTime.ParseExact(dateofdeathstr, "dd/MM/yyyy hh:mm:ss", Nothing)
                    dateofdeathstr = dateofdeath.ToString("MM-dd-yyyy")

                    Dim cmd1 As New OleDb.OleDbCommand
                    Dim dr1 As OleDb.OleDbDataReader
                    cmd1.Connection = cnn
                    cmd1.CommandText = "SELECT IIF(MAX(SLNO) IS NULL ,0,MAX(SLNO)) FROM DECEASECODE"
                    dr1 = cmd1.ExecuteReader
                    If dr1.Read = True Then
                        noofrecords = dr1(0)
                    Else
                        noofrecords = 0
                    End If
                    dr1.Close()

                    Dim cmd2 As New OleDb.OleDbCommand
                    cmd2.Connection = cnn
                    noofrecords = noofrecords + 1
                    cmd2.CommandText = "INSERT INTO DECEASECODE(SLNO,CID,DECEASEDATE)VALUES (" & noofrecords & ",'" & custid & "',#" & dateofdeathstr & "#)"
                    cmd2.ExecuteNonQuery()
                End While
                dr.Close()
                cnn.Close()
            End If
        Next
        oracle_conn.Close()
    End Sub
    Private Sub option820()     'Inserting from staff file
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim solid_int As Integer

        Dim custid As String
        Dim empno As String

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "STAFFCODE"

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = "SELECT SOLID,CID,EMPNO FROM STAFFLIST WHERE SOLID='" & solid_int & "'"
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()

                While dr.Read()
                    custid = ""
                    empno = ""

                    custid = dr(1).ToString()
                    empno = dr(2).ToString()

                    Dim cmd1 As New OleDb.OleDbCommand
                    Dim dr1 As OleDb.OleDbDataReader
                    cmd1.Connection = cnn
                    cmd1.CommandText = "SELECT COUNT(1) FROM PICKUP WHERE SLNO = 11  AND SUBSLNO = '" & empno.Trim & "'"
                    dr1 = cmd1.ExecuteReader
                    Dim temp As Integer = 0
                    If dr1.Read = True Then
                        temp = dr1(0)
                    End If
                    dr1.Close()
                    If temp = 1 Then
                        Dim cmd2 As New OleDb.OleDbCommand
                        cmd2.Connection = cnn
                        cmd2.CommandText = "UPDATE STAFFCODE SET STCODE ='" & empno & "' WHERE CID='" & custid & "'"
                        cmd2.ExecuteNonQuery()
                    End If

                End While
                dr.Close()
                cnn.Close()
            End If
        Next

        oracle_conn.Close()
        MsgBox("Process Completed")
    End Sub
    Private Sub option821()
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim solid_int As Integer

        Dim custid As String
        Dim categorygroup As String

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "CUSTCATEGORY"

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = "select solid,cid,decode (special, 'G' ,'GL','J','JA','X','EX','S','SW')special from specialcust where solid='" & solid_int & "'"
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()

                While dr.Read()
                    custid = ""
                    categorygroup = ""
                    custid = dr(1).ToString()
                    categorygroup = dr(2).ToString()
                    Dim cmd1 As New OleDb.OleDbCommand
                    Dim dr1 As OleDb.OleDbDataReader
                    cmd1.Connection = cnn
                    cmd1.CommandText = "SELECT COUNT(1) FROM  CUSTCATEGORY WHERE CID = '" & custid.Trim & "'"
                    dr1 = cmd1.ExecuteReader
                    Dim temp As Integer = 0
                    If dr1.Read = True Then
                        temp = dr1(0)
                    End If
                    dr1.Close()
                    If temp = 0 Then
                        Dim cmd2 As New OleDb.OleDbCommand
                        Dim dr2 As OleDb.OleDbDataReader
                        cmd2.Connection = cnn
                        cmd2.CommandText = "select cidname from cidmaster where cid='" & custid.Trim & "'"
                        dr2 = cmd2.ExecuteReader
                        Dim cidname As String = ""
                        If dr2.Read = True Then
                            cidname = dr2(0)
                        End If

                        If cidname <> "" Then
                            Dim cmd3 As New OleDb.OleDbCommand
                            cmd3.Connection = cnn
                            cmd3.CommandText = "INSERT INTO CUSTCATEGORY (CID,CIDNAME,CATEGORYGROUP) VALUES ('" & custid.Trim & "','" & cidname & "','" & categorygroup & "')"
                            cmd3.ExecuteNonQuery()
                        End If

                    End If

                End While
                dr.Close()
                cnn.Close()
            End If
        Next

        oracle_conn.Close()
        MsgBox("Process Completed")
    End Sub

    Private Sub option822()
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim solid_int As Integer

        Dim custid As String
        Dim religion As String
        Dim caste As String

        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "RELIGION"

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = "SELECT SOLID,CUSTID,DECODE (RELIGION,'H','1','M','2','C','3','B', '4''S','5')RELIGION,CASTE FROM RELIGION WHERE SOLID='" & solid_int & "'"
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()
                While dr.Read()
                    custid = ""
                    religion = ""
                    caste = ""
                    custid = dr(1).ToString()
                    religion = dr(2).ToString()
                    caste = dr(3).ToString()

                    Dim cmd1 As New OleDb.OleDbCommand
                    Dim dr1 As OleDb.OleDbDataReader
                    cmd1.Connection = cnn
                    cmd1.CommandText = "SELECT COUNT(1) FROM  RELIGION WHERE CUSTOMERID = '" & custid.Trim & "'"
                    dr1 = cmd1.ExecuteReader
                    Dim temp As Integer = 0
                    If dr1.Read = True Then
                        temp = dr1(0)
                    End If
                    dr1.Close()
                    If temp <> 0 Then
                        If religion <> "" Then
                            Dim cmd2 As New OleDb.OleDbCommand
                            cmd2.Connection = cnn
                            cmd2.CommandText = "UPDATE RELIGION SET RELIGIONCODE='" & religion & "' WHERE CUSTOMERID='" & custid & "'"
                            cmd2.ExecuteNonQuery()
                        End If
                        If caste <> "" Then
                            If caste.Trim = "C" Then
                                Dim cmd2 As New OleDb.OleDbCommand
                                cmd2.Connection = cnn
                                cmd2.CommandText = "UPDATE RELIGION SET CASTCODE='2' WHERE CUSTOMERID='" & custid & "'"
                                cmd2.ExecuteNonQuery()
                            End If
                            If caste.Trim = "T" Then
                                Dim cmd2 As New OleDb.OleDbCommand
                                cmd2.Connection = cnn
                                cmd2.CommandText = "UPDATE RELIGION SET CASTCODE='3' WHERE CUSTOMERID='" & custid & "'"
                                cmd2.ExecuteNonQuery()
                            End If
                            If caste.Trim = "O" Then
                                Dim cmd2 As New OleDb.OleDbCommand
                                cmd2.Connection = cnn
                                cmd2.CommandText = "UPDATE RELIGION SET CASTCODE='4' WHERE CUSTOMERID='" & custid & "'"
                                cmd2.ExecuteNonQuery()
                            End If
                        End If
                    End If
                    processmessage("Writing in file " & solid & "Record : " & custid)
                End While
                dr.Close()
                cnn.Close()
            End If
        Next
        MsgBox("Process Completed")
        oracle_conn.Close()
    End Sub
    Private Sub option823()
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim solid_int As Integer

        Dim custid As String
        Dim handicap As String


        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "CUSTCATEGORY"

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = "SELECT SOLID,CUSTID,DECODE (HANDICAP,'B','BP','H','HP')HANDICAP FROM HANDICAPPED WHERE SOLID='" & solid_int & "'"
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()
                While dr.Read()
                    custid = ""
                    handicap = ""

                    custid = dr(1).ToString()
                    handicap = dr(2).ToString()

                    Dim cmd1 As New OleDb.OleDbCommand
                    Dim dr1 As OleDb.OleDbDataReader
                    cmd1.Connection = cnn
                    cmd1.CommandText = "SELECT COUNT(1) FROM  CUSTCATEGORY WHERE CID = '" & custid.Trim & "'"
                    dr1 = cmd1.ExecuteReader
                    Dim temp As Integer = 0
                    If dr1.Read = True Then
                        temp = dr1(0)
                    End If
                    dr1.Close()
                    If temp = 1 Then
                        If handicap <> "" Then
                            Dim cmd2 As New OleDb.OleDbCommand
                            cmd2.Connection = cnn
                            cmd2.CommandText = "UPDATE CUSTCATEGORY SET  categorytype='" & handicap & "' WHERE CID='" & custid & "'"
                            cmd2.ExecuteNonQuery()
                        End If
                    Else
                        If handicap <> "" Then
                            Dim cmd2 As New OleDb.OleDbCommand
                            Dim dr2 As OleDb.OleDbDataReader
                            cmd2.Connection = cnn
                            cmd2.CommandText = "select cidname from cidmaster where cid='" & custid.Trim & "'"
                            dr2 = cmd2.ExecuteReader
                            Dim cidname As String = ""
                            If dr2.Read = True Then
                                cidname = dr2(0)
                            End If

                            If cidname <> "" Then
                                Dim cmd3 As New OleDb.OleDbCommand
                                cmd3.Connection = cnn
                                cmd3.CommandText = "INSERT INTO CUSTCATEGORY (CID,CIDNAME,CATEGORYTYPE) VALUES ('" & custid.Trim & "','" & cidname & "','" & handicap & "')"
                                cmd3.ExecuteNonQuery()
                            End If
                        End If
                    End If
                    processmessage("Writing in file :" & solid & "Record : " & custid)
                End While
                dr.Close()
                cnn.Close()
            End If
        Next
        oracle_conn.Close()
        MsgBox("Process over")

    End Sub

    Private Sub option824()
        Dim acess_db As String
        Dim access_db_path As String
        Dim accss_table As String
        Dim solid As String
        Dim solname() As String
        Dim sol As String
        Dim solid_int As Integer

        Dim acno As String
        Dim lpd_status As String


        access_db_path = InputBox("Enter the access database path", "Enter valu", "c:\du")
        acess_db = InputBox("Enter the access database path", "Enter valu", "NMGB.mdb")
        accss_table = "LPD"

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        For Each dir1 In Directory.GetDirectories(access_db_path)
            solname = dir1.Split("\")
            solid = solname(solname.Length - 1).Substring(0, 5)
            solname = solname(solname.Length - 1).Split("_")
            sol = solname(solname.Length - 1)
            solid_int = Val(solid)

            If File.Exists(dir1.ToString() & "\server\" & acess_db) Then

                Dim cnn As New OleDb.OleDbConnection
                cnn = New OleDb.OleDbConnection

                Dim strConnection As String
                strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & dir1.ToString() & "\server\" & acess_db
                cnn.ConnectionString = strConnection

                Try
                    If Not cnn.State = ConnectionState.Open Then
                        cnn.Open()
                    End If
                Catch ex As Exception
                    MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                End Try

                Dim sql As String
                sql = "SELECT SOLID,ACNO,LPD_STAT FROM LPD1 WHERE SOLID='" & solid_int & "'"
                Dim cmd As New OracleCommand(sql, oracle_conn)
                Dim dr As OracleDataReader = cmd.ExecuteReader()
                While dr.Read()
                    acno = ""
                    lpd_status = ""
                    acno = dr(1).ToString()
                    lpd_status = dr(2).ToString()
                    Dim cmd1 As New OleDb.OleDbCommand
                    Dim dr1 As OleDb.OleDbDataReader
                    cmd1.Connection = cnn
                    cmd1.CommandText = "SELECT CUSTNAME FROM ACMASTER  WHERE ACNO = '" & acno.Trim & "'"
                    dr1 = cmd1.ExecuteReader
                    Dim custname As String = ""

                    If dr1.Read = True Then
                        custname = dr1(0)
                    End If
                    dr1.Close()

                    If custname <> "" Then
                        If lpd_status = "E" Or lpd_status = "D" Then
                            lpd_status = "S"
                        End If
                        If lpd_status = "S" Or lpd_status = "R" Then
                            Dim cmd3 As New OleDb.OleDbCommand
                            cmd3.Connection = cnn
                            cmd3.CommandText = "INSERT INTO LPD (ACNO,CUSTNAME,LPDTYPE) VALUES ('" & acno.Trim & "','" & custname & "','" & lpd_status & "')"
                            cmd3.ExecuteNonQuery()
                        End If
                    End If
                    processmessage("Writing in file : " & solid & "Acno: " & acno)
                End While
                dr.Close()
                cnn.Close()
            End If
        Next
        oracle_conn.Close()
        MsgBox("Process Over")
    End Sub

    Public Sub compress(ByVal destination, ByVal directoryname, ByVal source)
        'Process.Start("""C:\Program Files\WinRAR\winrar.exe""", "a -afrar -m5 -ed -p -r -ep1  """ & destination & "\" & directoryname & ".rar" & """ """ & source)
        Process.Start("""C:\Program Files\WinRAR\winrar.exe""", "a -ap  """ & destination & "\" & directoryname & ".rar" & """ """ & source)
        System.Windows.Forms.Application.DoEvents()
        Thread.Sleep(1000)
    End Sub
    Public Sub mail_withattachment(ByVal destination, ByVal directoryname, ByVal source)
        Dim response As Integer
        response = MsgBox("Do you want to mail rar file?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm")
        If response = 6 Then
            Dim file1 As String = destination & "\" & directoryname & ".rar"
            'Dim file1 As String = destination & "\20140315_115312.rar"
            Dim oApp As Outlook._Application
            oApp = New Outlook.Application()
            Dim outlooksendfromaccount As String
            Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
            Dim dirs As String() = Directory.GetFiles("c:\temp")

            outlooksendfromaccount = "franklinkf@gmail.com"

            Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)

            newMail.To = "franklinkf@gmail.com"
            newMail.CC = "franklinkf@gmail.com"
            newMail.BCC = "franklinkf@gmail.com"
            newMail.Subject = "Back up of " & source & txtdate.Text
            newMail.HTMLBody = "<html><body><head><style>.normalandleft{TEXT-ALIGN: left; FONT-FAMILY: arial, helvetica, sans-serif; FONT-SIZE: 9pt; FONT-WEIGHT: normal;}</style></head><p class=normalandleft>Dear Sir,</p><p class=normalandleft>Enclosed please find the following files containing the back of " & source & " folder on " & txtdate.Text & ":</p><p class=normalandleft>This facility is only for keeping an online back up of your folders <br><br></p><p class=normalandleft></p><p class=normalandleft>Thanking you</p><p class=normalandleft><br><br><p class=normalandleft>Yours faithfully,</p><p class=normalandleft>MIS Team</p></body></html>"
            newMail.Attachments.Add(file1)
            newMail.SendUsingAccount = account
            newMail.Send()

            MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")

        End If


    End Sub
    Public Sub option825()       'Compress and mail folder
        Dim source As String
        Dim destination As String
        Dim directoryname As String
        directoryname = System.DateTime.Now
        directoryname = Format(System.DateTime.Now, "yyyy-MM-dd HH:mm:ss")
        directoryname = directoryname.Replace("-", "")
        directoryname = directoryname.Replace("\", "")
        directoryname = directoryname.Replace(":", "")
        directoryname = directoryname.Replace(" ", "_")
        source = InputBox("Enter the folder with full path which is to be compressed", "Enter value", Disk & ":\XXX")
        destination = InputBox("Enter the folder path where the compressed file is to be placed", "Enter value", Disk & ":\XXXXX")
        compress(destination, directoryname, source)
        Application.DoEvents()
        Thread.Sleep(1000)
        mail_withattachment(destination, directoryname, source)
        MsgBox("Process completed", MsgBoxStyle.Information)
    End Sub
    Public Sub Option826()   'Creating Differential backup
        Dim source As String
        Dim destination As String
        Dim directoryname As String
        Dim sourcepath() As String

        Dim backup_folder As String

        directoryname = System.DateTime.Now
        directoryname = Format(System.DateTime.Now, "yyyy-MM-dd HH:mm:ss")
        directoryname = directoryname.Replace("-", "")
        directoryname = directoryname.Replace("\", "")
        directoryname = directoryname.Replace(":", "")
        directoryname = directoryname.Replace(" ", "")
        source = InputBox("Enter the folder name with full path", "Enter value", Disk & ":\XXXXX")
        If source.Contains("\") Then
            sourcepath = source.Split("\")
            backup_folder = sourcepath(sourcepath.Length - 1)

        Else
            MsgBox("Enter Valid Source folder with full path (Eg: " & Disk & ":\Report pack")
        End If

        destination = InputBox("Enter the folder path where the backupkept", "Enter value", Disk & ":\XXXXX")

        Dim foldercount As Integer
        foldercount = Directory.GetDirectories(destination).Count

        If foldercount = 0 Then

            System.IO.Directory.CreateDirectory(destination & "\" & directoryname)
            destination = destination & "\" & directoryname
            copyfolder(source, destination, "Y", "N")
        Else
            ' For Each dir1 As String In Directory.GetDirectories(source)
            'dir1 = dir1.Replace(Disk & ":\", "")
            'check_latest_timestamp(dir1, destination)
            oracle_execute_non_query("ten", username, username, "truncate table c_misprint")
            oracle_execute_non_query("ten", username, username, "truncate table c_misadv")
            oracle_execute_non_query("ten", username, username, "truncate table c_misdep")

            directory_listing("Source", source, backup_folder, "Differential")
            directory_listing("Destination", destination, backup_folder, "Differential")

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            sql = "PKGMISTOOL2.FIND_LATEST_TIMESTAMP"
            Dim cmd1 As New OracleCommand(sql, conn)
            cmd1.CommandType = CommandType.StoredProcedure
            cmd1.Parameters.Add("EXTENSION", OracleDbType.Varchar2, 60, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd1.ExecuteNonQuery()

            conn.Close()
            conn.Dispose()

            createbackupofnewfiles(destination, backup_folder)
        End If
        MsgBox("Process completed", MsgBoxStyle.Information)
    End Sub

    Public Sub Option829()   'Creating Differential backup based on extension
        Dim source As String
        Dim destination As String
        Dim directoryname As String
        Dim sourcepath() As String
        Dim extensions As String
        Dim backup_folder As String

        directoryname = System.DateTime.Now
        directoryname = Format(System.DateTime.Now, "yyyy-MM-dd HH:mm:ss")
        directoryname = directoryname.Replace("-", "")
        directoryname = directoryname.Replace("\", "")
        directoryname = directoryname.Replace(":", "")
        directoryname = directoryname.Replace(" ", "")
        source = InputBox("Enter the folder name with full path", "Enter value", Disk & ":\XXXX")
        If source.Contains("\") Then
            sourcepath = source.Split("\")
            backup_folder = sourcepath(sourcepath.Length - 1)

        Else
            MsgBox("Enter Valid Source folder with full path (Eg: " & Disk & ":\Report pack")
        End If

        destination = InputBox("Enter the folder path where the backup is to be kept", "Enter value", Disk & ":\XXXX")
        extensions = InputBox("Enter the extensions of the files required to create backup", "Enter value", "ALL")
        Dim foldercount As Integer
        foldercount = Directory.GetDirectories(destination).Count
        'check_latest_timestamp(dir1, destination)
        oracle_execute_non_query("ten", username, username, "truncate table c_misprint")
        oracle_execute_non_query("ten", username, username, "truncate table c_misadv")
        oracle_execute_non_query("ten", username, username, "truncate table c_misdep")

        directory_listing("Source", source, backup_folder, "Differential")
        directory_listing("Destination", destination, backup_folder, "Differential")

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGMISTOOL2.FIND_LATEST_TIMESTAMP"
        Dim cmd1 As New OracleCommand(sql, conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.Add("EXTENSION", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = extensions
        cmd1.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()

        createbackupofnewfiles(destination, backup_folder)
        ' End If
        MsgBox("Process completed", MsgBoxStyle.Information)
    End Sub

    Public Sub Option830()   'Creating Mirror Image
        Dim source As String
        Dim destination As String
        Dim directoryname As String
        Dim sourcepath() As String
        Dim backup_folder As String
        Dim extension As String

        source = InputBox("Enter the folder name with full path", "Enter value", Disk & ":\CBS")
        If source.Contains("\") Then
            sourcepath = source.Split("\")
            'backup_folder = sourcepath(sourcepath.Length - 1)

        Else
            MsgBox("Enter Valid Source folder with full path (Eg: " & Disk & ":\Report pack")
        End If

        destination = InputBox("Enter the folder path where the backup is kept", "Enter value", Disk & ":\ONEDRIVE")
        extension = InputBox("Enter the extension", "Enter value", ".doc,.docx,.xls,.xlsx,.ppt,.pptx")
        Dim foldercount As Integer
        foldercount = Directory.GetDirectories(destination).Count

        oracle_execute_non_query("ten", username, username, "truncate table c_misprint")
        oracle_execute_non_query("ten", username, username, "truncate table c_misadv")
        oracle_execute_non_query("ten", username, username, "truncate table c_misdep")

        directory_listing("Source", source, backup_folder, "Mirror")
        directory_listing("Destination", destination, backup_folder, "Mirror")

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "PKGMISTOOL2.COMPARE_SOURCE_DESTINATION"
        Dim cmd1 As New OracleCommand(sql, conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.Add("EXTENSION", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = extension
        cmd1.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()
        mirrorimage_source_destination(source, destination)

        MsgBox("Process completed", MsgBoxStyle.Information)
    End Sub

    Private Sub option831()  'Generating CIDMaster File From dump

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        Dim solid As String = InputBox("Enter Solid", "Enter Value", "00077")

        processmessage("Executing Package")
        sql = "PKGNMGB_CID_EXTRACTION.PROCESS"
        Dim cmd5 As New OracleCommand(sql, oracle_conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("SOLID", OracleDbType.Varchar2, 50, Nothing, ParameterDirection.Input).Value = Trim(solid)
        cmd5.ExecuteNonQuery()

        Dim sql1 As String
        sql1 = "SELECT  REPORTDATA FROM C_MISPRINT WHERE SOLID ='" & solid & "' ORDER BY SERIALNO"
        Dim cmd1 As New OracleCommand(sql1, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()

        Dim sw As StreamWriter = New StreamWriter("C:\du\" & solid & "_CIDMASTER.txt")
        While dr1.Read
            Dim linedata As String
            linedata = dr1(0)
            sw.WriteLine(linedata)
        End While
        dr1.Close()
        sw.Close()

        oracle_conn.Close()
        MsgBox("File generated Successfully")
    End Sub
    Private Sub option832()  'Create text files in a loop

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()
        Dim solid As String
        Dim datastring As String

        'sql = "SELECT DISTINCT BR_NO FROM LOAN_COMPL"
        'Dim cmd1 As New OracleCommand(sql, oracle_conn)
        'Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
        'While dr1.Read()
        '    solid = dr1.Item("BR_NO").ToString
        '    sql = "SELECT SUBSTR(KEY_1,10,10)||CHECKDIGIT||'|'||PRODUCTCODE||'|'||TO_CHAR(APPRV_DATE,'DD-MM-YYYY')||'|'||APP_AMT||'|'||LOAN_TRM||'|'||NVL(INTEREST,0)||'|'||NVL(STORE_RATE,0)||'|'||DECODE(REPAY_OPTION,'1','1 EQUATED INSTMT','2','2 PRINC EQ DISTR','3','3 PROJECTED INT','4','4 NEGOTIATED','5','5 STAFF LOAN','0 NO SCHEDULE')||'|'||DECODE(REPAY_FREQ,'01','01 MONTHLY', '03','03 QUARTERELY','06','06 HALF YEARLY','12','12 YEARLY','98','98 END OF TERM','00 INVALID')||'|'||LOAN_BAL ABCD FROM LOAN_COMPL WHERE BR_NO='" & solid & "' ORDER BY PRODUCTCODE,KEY_1"
        '    Dim cmd2 As New OracleCommand(sql, oracle_conn)
        '    Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
        '    Dim sw2 As StreamWriter = New StreamWriter("C:\DU\" & solid & "_LRS.txt")
        '    processmessage("Creating file of " & solid & " branch")
        '    While dr2.Read
        '        datastring = dr2.Item("ABCD").ToString
        '        sw2.WriteLine(datastring)
        '    End While
        '    sw2.Close()
        '    dr2.Close()
        'End While
        'dr1.Close()
        'oracle_conn.Close()
        'MsgBox("Files generated Successfully")

        sql = "SELECT DISTINCT HOME_BRANCH_NO FROM CUSTID1"
        Dim cmd1 As New OracleCommand(sql, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
        While dr1.Read()
            solid = dr1.Item("HOME_BRANCH_NO").ToString

            sql = "PKGNMGB_CID_EXTRACTION.PROCESS"
            Dim cmd5 As New OracleCommand(sql, oracle_conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.Parameters.Add("GSOLID", OracleDbType.Varchar2, Nothing, ParameterDirection.Input).Value = solid
            cmd5.ExecuteNonQuery()

            sql = "SELECT REPORTDATA FROM C_MISPRINT"
            Dim cmd2 As New OracleCommand(sql, oracle_conn)
            Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
            Dim sw2 As StreamWriter = New StreamWriter("C:\DU\" & solid & "_CIDMASTER.txt")
            processmessage("Creating file of " & solid & " branch")
            While dr2.Read
                datastring = dr2.Item("REPORTDATA").ToString
                sw2.WriteLine(datastring)
            End While
            sw2.Close()
            dr2.Close()
        End While
        dr1.Close()
        oracle_conn.Close()
        MsgBox("Files generated Successfully")
    End Sub

    Private Sub option833()  'Create text files in a loop

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        Dim cnn As New OleDb.OleDbConnection
        cnn = New OleDb.OleDbConnection

        Dim strConnection As String
        strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & "C:\nmgb\serverdatabasepath\nmgb.mdb"
        cnn.ConnectionString = strConnection

        Try
            If Not cnn.State = ConnectionState.Open Then
                cnn.Open()
            End If
        Catch ex As Exception
            MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
        End Try

        Dim sql1 As String
        sql1 = "SELECT  REPORTDATA FROM C_MISPRINT WHERE SERIALNO=5"
        Dim cmd1 As New OracleCommand(sql1, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
        Dim cid_var As String
        Dim address3 As String
        Dim name As String
        Dim linedata As String

        While dr1.Read()
            linedata = ""
            linedata = dr1("reportdata").ToString()
            cid_var = ""
            address3 = ""
            name = ""
            cid_var = linedata.Substring(0, 20).Trim
            name = linedata.Substring(20, 50).Trim
            address3 = linedata.Substring(70, 35).Trim
            Dim cmd11 As New OleDb.OleDbCommand
            cmd11.CommandText = "UPDATE CIDMASTER SET ADDRESS3= ADDRESS3 +'" & address3 & "' WHERE CID='" & cid_var & "'" 'and CIDMASTER.APPENDNUMBER=TRUE"
            cmd11.Connection = cnn
            cmd11.ExecuteNonQuery()
            processmessage("Writing record  :" & cid_var)
        End While
        dr1.Close()
        MsgBox("Completed Successfully")
        cnn.Close()
        oracle_conn.Close()
    End Sub


    Private Sub option601()   'eNMGB Migration - Create Branch Data
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        Dim import_user As String
        Dim solid As String
        Dim uconfirm As String

        solid = InputBox("Enter Solid", "Enter Value")
        import_user = InputBox("Username in which data to be imported", "Enter Value", "")
        solid = solid.PadLeft(5, "0")

        If Not Directory.Exists(Disk & ":\DUMP\BR_" & solid) Then
            MsgBox(Disk & ":\DUMP\BR_" & solid & " folderpath not found. Create the folder, place dump and try again")
            Exit Sub
        End If

        If Not File.Exists(Disk & ":\DUMP\BR_" & solid & "\B" & solid & ".dmp") Then
            MsgBox(Disk & ":\DUMP\BR_" & solid & "\B" & solid & ".dmp file not found. Place dump and try again")
            Exit Sub
        End If

        If Not File.Exists(Disk & ":\DUMP\BR_" & solid & "\T" & solid & ".dmp") Then
            MsgBox(Disk & ":\DUMP\BR_" & solid & "\T" & solid & ".dmp file not found. Place dump and try again")
            Exit Sub
        End If

        '' drop and create user

        Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
        sw0.WriteLine("@echo off")
        sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
        sw0.Close()

        Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
        sw.WriteLine("connect sys/sys@ten as sysdba;")
        sw.WriteLine("drop user " & import_user & " cascade;")
        sw.WriteLine("create user  " & import_user & " identified by  " & import_user & " default tablespace domain quota unlimited on domain;")
        sw.WriteLine("grant connect,resource,dba to  " & import_user & " ;")
        sw.WriteLine("alter user  " & import_user & " temporary tablespace tempdomain;")
        sw.WriteLine("grant create role to  " & import_user & " ;")
        sw.WriteLine("grant execute on XMLDOM to  " & import_user & " ;")
        sw.WriteLine("grant execute on XMLPARSER to  " & import_user & " ;")
        sw.Close()

        Process.Start(Disk & ":\dump\script\create_user.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Importing main data

        Dim sw1 As StreamWriter = New StreamWriter(Disk & ":\dump\script\import_user_br.bat")
        sw1.WriteLine("@echo off")
        sw1.WriteLine("imp " & import_user & "/" & import_user & "@ten file=" & Disk & ":\DUMP\BR_" & solid & "\B" & solid & ".dmp  full=yes")
        sw1.Close()
        Process.Start(Disk & ":\dump\script\import_user_br.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Renaming tran tables

        Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\dump\script\rename_table.sql")
        sw2.WriteLine("rename mig_tran to mig_tran1;")
        sw2.WriteLine("rename mig_tran_narration to mig_tran_narration1;")
        sw2.Close()

        Dim sw3 As StreamWriter = New StreamWriter(Disk & ":\dump\script\rename_table.bat")
        sw3.WriteLine("@echo off")
        sw3.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\script\rename_table.sql /nolog ")
        sw3.Close()

        Process.Start(Disk & ":\dump\script\rename_table.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Importing tran table

        Dim sw4 As StreamWriter = New StreamWriter(Disk & ":\dump\script\import_user_tr.bat")
        sw4.WriteLine("@echo off")
        sw4.WriteLine("imp " & import_user & "/" & import_user & "@ten file=" & Disk & ":\DUMP\BR_" & solid & "\T" & solid & ".dmp  full=yes")
        sw4.Close()
        Process.Start(Disk & ":\dump\script\import_user_tr.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Moving tran data

        Dim sw5 As StreamWriter = New StreamWriter(Disk & ":\dump\script\move_data.sql")
        sw5.WriteLine("insert into mig_tran select * from mig_tran1;")
        sw5.WriteLine("insert into mig_tran_narration select * from mig_tran_narration1;")
        sw5.WriteLine("drop table mig_tran1;")
        sw5.WriteLine("drop table mig_tran_narration1;")
        sw5.Close()

        Dim sw6 As StreamWriter = New StreamWriter(Disk & ":\dump\script\move_data.bat")
        sw6.WriteLine("@echo off")
        sw6.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\script\move_data.sql /nolog ")
        sw6.Close()

        Process.Start(Disk & ":\dump\script\move_data.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Create/alter tables 

        Dim sw7 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_table.bat")
        sw7.WriteLine("@echo off")
        sw7.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\Table1.sql /nolog ")
        sw7.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\Table2.sql /nolog ")
        sw7.Close()

        Process.Start(Disk & ":\dump\script\create_table.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Provide grant on synonym

        Dim sw8 As StreamWriter = New StreamWriter(Disk & ":\dump\script\mig_grant.sql")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON ADVANCE_MAPPING TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON ADVANCE_PURPOSE TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON BRM001 TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON CEDGE_PICKUP TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON CID TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON DEPOSIT_MAPPING TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON GL_MAPPING TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON INTEREST_CCOD TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON INTEREST_CUSTOM TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON INTEREST_DEFAULT TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON NNTM TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON PICKUP_MAPPING TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON REPORT_INDEX TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON SOL_MAPPING TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON SUB_GL_CODE TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON NND_PROVISION TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON ZENITH_TBA_GLCC_INT_TDS TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON FUF_INDEX TO " & import_user & ";")
        sw8.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON REPORT_COMPARISON_INDEX TO " & import_user & ";")
        sw8.Close()

        Dim sw9 As StreamWriter = New StreamWriter(Disk & ":\dump\script\cbs_grant.sql")
        sw9.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON C_MISLINKAGE TO " & import_user & ";")
        sw9.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON GSP TO " & import_user & ";")
        sw9.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON LSP TO " & import_user & ";")
        sw9.Close()

        Dim sw110 As StreamWriter = New StreamWriter(Disk & ":\dump\script\grant.bat")
        sw110.WriteLine("@echo off")
        sw110.WriteLine("sqlplus mig/mig@ten @" & Disk & ":\dump\script\mig_grant.sql /nolog")
        sw110.WriteLine("sqlplus cbs/cbs@ten @" & Disk & ":\dump\script\cbs_grant.sql /nolog")
        sw110.Close()

        Process.Start(Disk & ":\dump\script\grant.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Create package 

        Dim sw11 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_package.bat")
        sw11.WriteLine("@echo off")
        sw11.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\Package.sql /nolog ")
        sw11.Close()

        Process.Start(Disk & ":\dump\script\create_package.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Run scripts

        Dim sw12 As StreamWriter = New StreamWriter(Disk & ":\dump\script\run_scripts.bat")
        sw12.WriteLine("@echo off")
        sw12.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\Update.sql /nolog ")
        sw12.Close()

        Process.Start(Disk & ":\dump\script\run_scripts.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub
    Private Sub option612()   'eNMGB Migration - Create Branch Data

        Dim uname As String
        Dim uconfirm As String
        Dim startno = InputBox("Enter MIG Start user name (Number only)", "Start", "1")
        Dim endno = InputBox("Enter MIG End user name (Number only)", "End", "50")
        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\batch_update.bat")
        sw0.WriteLine("@echo off")
        sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\batch_update.sql")
        sw0.Close()

        Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\batch_update.sql")
        sql = "SELECT USERNAME FROM (SELECT USERNAME FROM DBA_USERS WHERE SUBSTR(USERNAME,1,3) = 'MIG') WHERE TO_NUMBER(SUBSTR(USERNAME,4,3)) BETWEEN " & startno & " AND " & endno & " AND EXISTS (SELECT 1 FROM DBA_TABLES WHERE OWNER = USERNAME AND TABLE_NAME = 'MIG_SUMMARY') ORDER BY USERNAME"
        Dim cmd As New OracleCommand(sql, oracle_conn)
        Dim dr As OracleDataReader = cmd.ExecuteReader()
        sw.WriteLine("connect mig/mig@ten;")
        sw.WriteLine("ALTER TABLE SOL_MAPPING ADD (MIGRATION_DATE DATE);")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40401;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40402;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40423;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40428;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40439;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40444;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40446;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40452;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40455;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40456;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40465;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40472;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40474;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40480;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40496;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40497;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40499;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40508;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40512;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40523;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40528;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40530;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40532;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40534;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40535;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40537;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40538;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40544;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40545;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40546;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40552;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40561;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40567;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40568;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40569;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40570;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40580;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40586;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40592;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40593;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40597;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40600;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40601;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40602;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40608;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40610;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40612;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40614;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40618;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='05-jul-2014' WHERE FIN_SOLID=40619;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40407;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40422;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40425;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40426;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40431;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40434;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40440;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40441;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40443;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40445;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40450;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40460;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40463;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40467;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40468;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40482;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40487;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40489;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40491;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40498;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40501;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40503;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40505;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40513;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40514;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40533;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40536;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40540;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40542;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40547;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40548;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40551;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40553;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40554;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40557;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40559;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40560;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40562;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40566;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40575;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40576;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40579;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40581;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40584;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40588;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40590;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40591;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40594;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40604;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40605;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='12-jul-2014' WHERE FIN_SOLID=40607;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40403;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40404;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40405;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40408;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40409;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40410;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40411;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40414;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40418;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40419;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40424;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40427;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40433;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40436;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40447;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40448;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40453;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40457;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40461;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40462;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40464;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40466;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40476;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40478;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40479;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40481;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40484;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40486;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40488;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40490;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40493;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40494;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40495;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40502;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40504;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40507;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40516;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40518;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40520;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40521;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40522;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40527;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40531;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40539;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40541;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40543;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40549;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40550;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40555;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40556;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40558;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40582;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40587;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40599;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40603;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40611;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-jul-2014' WHERE FIN_SOLID=40656;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='22-june-2014' WHERE FIN_SOLID=40421;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='22-june-2014' WHERE FIN_SOLID=40477;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='22-june-2014' WHERE FIN_SOLID=40517;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-JUL-2014' WHERE FIN_SOLID=40392;")
        sw.WriteLine("UPDATE SOL_MAPPING SET MIGRATION_DATE='19-JUL-2014' WHERE FIN_SOLID=40394;")
        While dr.Read
            uname = dr.Item("USERNAME").ToString.Trim
            sw.WriteLine("connect mig/mig@ten;")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON ADVANCE_MAPPING TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON ADVANCE_PURPOSE TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON BRM001 TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON CEDGE_PICKUP TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON CID TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON DEPOSIT_MAPPING TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON GL_MAPPING TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON INTEREST_CCOD TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON INTEREST_CUSTOM TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON INTEREST_DEFAULT TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON NNTM TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON PICKUP_MAPPING TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON REPORT_INDEX TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON SOL_MAPPING TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON SUB_GL_CODE TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON NND_PROVISION TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON ZENITH_TBA_GLCC_INT_TDS TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON FUF_INDEX TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON REPORT_COMPARISON_INDEX TO " & uname & ";")
            sw.WriteLine("connect cbs/cbs@ten;")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON C_MISLINKAGE TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON GSP TO " & uname & ";")
            sw.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON LSP TO " & uname & ";")
            sw.WriteLine("connect " & uname & "/" & uname & "@ten;")
            sw.WriteLine("CREATE INDEX C_MISADV_IDX1 ON C_MISADV (ACNO,DATE1);")
            sw.WriteLine("connect " & uname & "/" & uname & "@ten;")
            sw.WriteLine("@" & Disk & ":\dump\static\Package.sql /nolog ")
        End While
        sw.Close()
        dr.Close()

        Process.Start(Disk & ":\dump\script\batch_update.bat")
        oracle_conn.Close()

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        MsgBox("Process completed.  Click OK to view the details of uncompiled packages", MsgBoxStyle.Information, "Process Completed")
        sql = "SELECT OWNER||' - '||OBJECT_NAME REPORTDATA FROM ALL_OBJECTS WHERE OBJECT_TYPE = 'PACKAGE' AND STATUS <> 'VALID' AND OWNER IN (SELECT USERNAME FROM (SELECT USERNAME FROM DBA_USERS WHERE SUBSTR(USERNAME,1,3) = 'MIG') WHERE TO_NUMBER(SUBSTR(USERNAME,4,3)) BETWEEN " & startno & " AND " & endno & " AND EXISTS (SELECT 1 FROM DBA_TABLES WHERE OWNER = USERNAME AND TABLE_NAME = 'MIG_SUMMARY'))"
        display_in_File(sql, "C:\du\aa.txt")
        Process.Start("C:\du\aa.txt")

    End Sub
    Private Sub option613()   'eNMGB Migration - Create backup of live users

        Dim uname As String
        Dim uconfirm As String
        Dim startno = InputBox("Enter MIG Start user name (Number only)", "Start", "1")
        Dim endno = InputBox("Enter MIG End user name (Number only)", "End", "50")

        Dim oracle_cnn_string As String = "Data Source=ten;User Id=cbs;Password=cbs;"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\613_1.bat")
        sw0.WriteLine("@echo off")
        sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\613_1.sql")
        sw0.Close()

        Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\613_1.sql")
        sw.WriteLine("connect sys/sys@ten as sysdba;")
        sw.WriteLine("GRANT SELECT ON DBA_USERS TO CBS;")
        sw.WriteLine("GRANT SELECT ON DBA_TABLES TO CBS;")
        sw.Close()

        Process.Start(Disk & ":\dump\script\613_1.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        Dim sw1 As StreamWriter = New StreamWriter(Disk & ":\dump\script\613_2.bat")
        sw1.WriteLine("@echo off")
        sw1.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\613_2.sql")
        sw1.Close()

        Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\dump\script\613_2.sql")
        sw2.WriteLine("connect mig/mig@ten;")
        sw2.WriteLine("GRANT SELECT, INSERT, UPDATE, DELETE ON SOL_MAPPING TO CBS;")
        sw2.Close()

        Process.Start(Disk & ":\dump\script\613_2.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        Dim sw3 As StreamWriter = New StreamWriter(Disk & ":\dump\script\613_3.bat")
        sw3.WriteLine("@echo off")
        sw3.WriteLine("sqlplus cbs/cbs@ten @" & Disk & ":\dump\static\pkgtemp1.sql /nolog ")
        sw3.Close()

        Process.Start(Disk & ":\dump\script\613_3.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        'sql = "PKGTEMP1.MISC1"
        'Dim cmd1 As New OracleCommand(sql, oracle_conn)
        'cmd1.CommandType = CommandType.StoredProcedure
        'cmd1.ExecuteNonQuery()



        Dim cedge_solid As String
        Dim finacle_solid As String

        sql = "SELECT USERNAME FROM DBA_USERS WHERE SUBSTR(USERNAME,1,3) = 'MIG' AND EXISTS (SELECT 1 FROM DBA_TABLES WHERE OWNER = USERNAME AND TABLE_NAME = 'MIG_SUMMARY') AND TO_NUMBER(SUBSTR(USERNAME,4,3)) BETWEEN " & startno & " AND " & endno
        Dim cmd As New OracleCommand(sql, oracle_conn)
        Dim dr As OracleDataReader = cmd.ExecuteReader()
        While dr.Read

            uname = dr.Item("USERNAME").ToString.Trim
            sql = "SELECT BRANCH_NO FROM " & uname & ".MIG_SUMMARY WHERE ROWNUM < 2"
            Dim cmd11 As New OracleCommand(sql, oracle_conn)
            Dim dr1 As OracleDataReader = cmd11.ExecuteReader()
            While dr1.Read

                cedge_solid = dr1.Item("BRANCH_NO").ToString.Trim

                sql = "SELECT FIN_SOLID FROM SM WHERE CED_SOLID = LTRIM('" & cedge_solid & "','0')"
                Dim cmd12 As New OracleCommand(sql, oracle_conn)
                Dim dr12 As OracleDataReader = cmd12.ExecuteReader()
                While dr12.Read

                    finacle_solid = dr12.Item("FIN_SOLID").ToString.Trim

                End While
                dr12.Close()

                processmessage("Exporting data of " & uname & "user")
                Dim sw17 As StreamWriter = New StreamWriter(Disk & ":\dump\script\export_user_" & uname & ".bat")
                sw17.WriteLine("@echo off")
                'sw17.WriteLine("exp " & uname & "/" & uname & "@TEN file=" & Disk & ":\" & bcode & ".dmp log=" & Disk & ":\" & bcode & ".log")
                sw17.WriteLine("exp " & uname & "/" & uname & "@TEN file=" & Disk & ":\" & finacle_solid & "_" & cedge_solid & "_" & uname & ".dmp /nolog")
                sw17.Close()

                Process.Start(Disk & ":\dump\script\export_user_" & uname & ".bat")

                uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                If uconfirm <> "Y" Then
                    MsgBox("Exiting application")
                    Exit Sub
                End If

            End While
            dr1.Close()

        End While
        sw.Close()
        dr.Close()

        MsgBox("Process completed", MsgBoxStyle.Information, "Process Completed")

    End Sub
    Private Sub option614()   'eNMGB Migration - Import Users

        Dim dumppath As String = InputBox("Enter path in which is backups are placed", "Enter Value", "")
        Dim usernameprefix As String = InputBox("Oracle username - Prefix to be added to file name for username", "Enter Value", "NIL")
        Dim filelength As String = InputBox("Oracle username - No. of characters of file name to be reckoned for username", "Enter Value", "ALL")
        Dim uconfirm As String
        If dumppath = "" Then
            MsgBox("No data entered.  Exiting application")
            Exit Sub
        End If

        Dim tempcount As Integer
        Dim processrunning As Boolean
        Dim myProcesses() As Process
        Dim myProcess As Process

        Dim dirs As String() = Directory.GetFiles(dumppath, "*.dmp")
        Dim dir As String
        Dim totalfiles As Integer
        Dim filename As String
        Dim uname As String

        totalfiles = dirs.Length

        If totalfiles = 0 Then
            MsgBox("No .dmp exists in the folder " & dumppath, MsgBoxStyle.Critical, "Error")
        Else
            For Each dir In dirs
                tempcount = tempcount + 1
                filename = GetFileName(dir)
                If usernameprefix <> "NIL" Then
                    uname = usernameprefix
                End If
                If filelength = "ALL" Then
                    uname = uname & filename
                Else
                    uname = uname & filename.Substring(0, filelength)
                End If

                Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
                sw0.WriteLine("@echo off")
                sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
                sw0.Close()

                Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
                sw.WriteLine("connect sys/sys@ten as sysdba;")
                sw.WriteLine("drop user " & uname & " cascade;")
                sw.WriteLine("create user  " & uname & " identified by  " & uname & " default tablespace domain quota unlimited on domain;")
                sw.WriteLine("grant connect,resource,dba to  " & uname & " ;")
                sw.WriteLine("alter user  " & uname & " temporary tablespace tempdomain;")
                sw.WriteLine("grant create role to  " & uname & " ;")
                sw.WriteLine("grant execute on XMLDOM to  " & uname & " ;")
                sw.WriteLine("grant execute on XMLPARSER to  " & uname & " ;")
                sw.WriteLine("exit")
                sw.Close()

                Process.Start(Disk & ":\dump\script\create_user.bat")

                'uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                'If uconfirm <> "Y" Then
                '    MsgBox("Exiting application")
                '    Exit Sub
                'End If

                processrunning = True
                While processrunning
                    tempcount = 0
                    myProcesses = Process.GetProcesses()
                    For Each myProcess In myProcesses
                        If UCase(myProcess.ProcessName) = "CMD" Then
                            tempcount = 1
                        End If
                    Next
                    If tempcount = 0 Then
                        processrunning = False
                    End If
                    If processrunning = True Then
                        Thread.Sleep(2000)
                    End If
                End While

                Dim sw1 As StreamWriter = New StreamWriter(Disk & ":\dump\script\import_user_br.bat")
                sw1.WriteLine("@echo off")
                sw1.WriteLine("imp " & uname & "/" & uname & "@ten file=" & dir & " full=yes")
                sw1.Close()
                Process.Start(Disk & ":\dump\script\import_user_br.bat")

                processrunning = True
                While processrunning
                    tempcount = 0
                    myProcesses = Process.GetProcesses()
                    For Each myProcess In myProcesses
                        If UCase(myProcess.ProcessName) = "CMD" Then
                            tempcount = 1
                        End If
                    Next
                    If tempcount = 0 Then
                        processrunning = False
                    End If
                    If processrunning = True Then
                        Thread.Sleep(2000)
                    End If
                End While

                'uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                'If uconfirm <> "Y" Then
                '    MsgBox("Exiting application")
                '    Exit Sub
                'End If

            Next
        End If

        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub
    Private Sub option615()   'eNMGB Migration - Data from users

        Dim uname As String
        Dim tempcount As Integer
        Dim processrunning As Boolean
        Dim myProcesses() As Process
        Dim myProcess As Process
        Dim slno As Integer = 0

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "SELECT USERNAME FROM DBA_USERS WHERE SUBSTR(USERNAME,1,2) = 'M4' ORDER BY USERNAME"
        'sql = "SELECT USERNAME FROM DBA_USERS WHERE SUBSTR(USERNAME,1,6) = 'M40600' ORDER BY USERNAME"
        Dim cmd4 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd4.ExecuteReader()
        While dr.Read()
            uname = dr.Item("USERNAME").ToString
            slno = slno + 1
            processmessage("Processing User No - " & slno & " (" & uname & ")")


            Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
            sw0.WriteLine("@echo off")
            sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
            sw0.Close()

            Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
            sw.WriteLine("connect " & uname & "/" & uname & "@ten;")
            'sw.WriteLine("DELETE FROM C_MISADV;")
            'sw.WriteLine("COMMIT;")
            'sw.WriteLine("INSERT INTO C_MISADV (SOLID,ACNO,TEXT1,NUMBER1) SELECT F_SOLID,F_ACNO,C_ACNO,ROUND(TOVERDUE)*-1 FROM FUF_LRS WHERE TOVERDUE < 0;")
            'sw.WriteLine("COMMIT;")
            'sw.WriteLine("UPDATE C_MISADV SET TEXT2 = (SELECT ACID FROM FIN_GAM WHERE FORACID = ACNO);")
            'sw.WriteLine("COMMIT;")
            'sw.WriteLine("INSERT INTO CBS.C_MISPRINT (SOLID,REPORTDATA) SELECT SOLID,SOLID||'|'||TEXT2||'|'||ACNO||'|'||TEXT1||'|'||NUMBER1||'|||' FROM C_MISADV;")
            'sw.WriteLine("COMMIT;")
            'sw.WriteLine("EXIT")
            'sw.WriteLine("INSERT INTO CBS.C_MISPRINT (SOLID,REPORTDATA) SELECT F_SOLID,F_SOLID||'|'||F_ACNO||'|'||TOVERDUE||'|'||POVERDUE||'|'||IBAL3 FROM FUF_LRS WHERE POVERDUE < 0;")
            'sw.WriteLine("COMMIT;")
            'sw.WriteLine("DELETE FROM C_MISADV;")
            'sw.WriteLine("INSERT INTO C_MISADV (SOLID,TEXT1,TEXT2,TEXT3,TEXT4,DATE1,NUMBER1,NUMBER2,NUMBER3,NUMBER4,NUMBER5,NUMBER6,NUMBER7) SELECT F_SOLID,PRODUCT_CODE,SCHEME_CODE,F_ACNO,C_ACNO,FIRST_INSTALLMENT_DATE,LOANAMOUNT,DISBURSEMENT,PBAL3,IBAL3,BALOS,POVERDUE,TOVERDUE FROM FUF_LRS WHERE (TOVERDUE < 0 OR POVERDUE < 0);")
            'sw.WriteLine("UPDATE C_MISADV SET TEXT5 = (SELECT ACID FROM FIN_GAM WHERE FORACID = TEXT3);")
            'sw.WriteLine("INSERT INTO CBS.C_MISADV (SOLID,TEXT1,TEXT2,TEXT3,TEXT4,DATE1,NUMBER1,NUMBER2,NUMBER3,NUMBER4,NUMBER5,NUMBER6,NUMBER7,TEXT5) SELECT SOLID,TEXT1,TEXT2,TEXT3,TEXT4,DATE1,NUMBER1,NUMBER2,NUMBER3,NUMBER4,NUMBER5,NUMBER6,NUMBER7,TEXT5 FROM C_MISADV;")
            sw.WriteLine("INSERT INTO THREE.MIG_ONRF SELECT BRANCH_NO,BAC,OLD_CUST_ACCT_NO FROM MIG_ONRF WHERE NO_TYPE = 'A';")
            sw.WriteLine("INSERT INTO THREE.MIG_ACC_NO SELECT BRANCH_NO,SCHEME_TYPE,SCHEME_CODE,PRODUCT_CODE,AC_NO_CD,FORACID,ACCT_CATEGORY FROM MIG_ACC_NO;")
            sw.WriteLine("EXIT")
            sw.Close()

            Process.Start(Disk & ":\dump\script\create_user.bat")

            processrunning = True
            While processrunning
                tempcount = 0
                myProcesses = Process.GetProcesses()
                For Each myProcess In myProcesses
                    If UCase(myProcess.ProcessName) = "CMD" Then
                        tempcount = 1
                    End If
                Next
                If tempcount = 0 Then
                    processrunning = False
                End If
                If processrunning = True Then
                    Thread.Sleep(2000)
                End If
            End While

        End While
        dr.Close()
        conn.Close()

        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Private Sub option609()   'eNMGB Migration - Split CEDGE Dump
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        Dim import_user As String
        Dim uconfirm As String
        Dim dumpname As String

        import_user = InputBox("Username in which data to be imported", "Enter Value", "TEMP")
        dumpname = InputBox("Enter CEDGE Dump Name (with path)", "Enter Value", "")

        '' drop and create user

        Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\split_dump_create_user.bat")
        sw0.WriteLine("@echo off")
        sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\static\split_dump_create_user.sql")
        sw0.Close()
        Process.Start(Disk & ":\dump\script\split_dump_create_user.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Create Tables

        Dim sw6 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_tables.bat")
        sw6.WriteLine("@echo off")
        sw6.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\split_dump_create_tables.sql /nolog ")
        sw6.Close()

        Process.Start(Disk & ":\dump\script\create_tables.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Import Dump

        Dim sw4 As StreamWriter = New StreamWriter(Disk & ":\dump\script\import_cedge_dump.bat")
        sw4.WriteLine("@echo off")
        sw4.WriteLine("imp " & import_user & "/" & import_user & "@ten file=" & dumpname & "  full=yes")
        sw4.Close()
        Process.Start(Disk & ":\dump\script\import_cedge_dump.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Move Data

        Dim sw16 As StreamWriter = New StreamWriter(Disk & ":\dump\script\update_package.bat")
        sw16.WriteLine("@echo off")
        sw16.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\split_dump_move_data.sql /nolog ")
        sw16.Close()

        Process.Start(Disk & ":\dump\script\update_package.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & import_user & ";Password= " & import_user & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        Dim sql5 As String
        Dim uname As String
        Dim bcode As String
        sql5 = "SELECT 'B'||BRANCHCODE BRCODE,USERNAME FROM (SELECT DISTINCT BRANCHCODE,USERNAME FROM ABCD) ORDER BY USERNAME"
        Dim cmd5 As New OracleCommand(sql5, oracle_conn)
        Dim dr5 As OracleDataReader = cmd5.ExecuteReader()
        While dr5.Read

            uname = dr5.Item("USERNAME").ToString.Trim
            bcode = dr5.Item("BRCODE").ToString.Trim
            processmessage("Exporting data of " & uname & "user")
            Dim sw17 As StreamWriter = New StreamWriter(Disk & ":\dump\script\export_user_" & uname & ".bat")
            sw17.WriteLine("@echo off")
            sw17.WriteLine("exp " & uname & "/" & uname & "@TEN file=" & Disk & ":\" & bcode & ".dmp log=" & Disk & ":\" & bcode & ".log")
            sw17.Close()

            Process.Start(Disk & ":\dump\script\export_user_" & uname & ".bat")

            uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
            If uconfirm <> "Y" Then
                MsgBox("Exiting application")
                Exit Sub
            End If

        End While
        dr5.Close()

        processmessage("")

        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub
    Private Sub option611()   'eNMGB Migration - Create NPA Upload Files

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim tempcount As Integer
        processmessage("Checking prerequisites")
        sql = "SELECT COUNT(1) COUNT FROM FIN_GAM"
        Dim cmd1 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd1.ExecuteReader()
        While dr.Read
            tempcount = dr.Item("COUNT").ToString.Trim
        End While
        dr.Close()

        If tempcount = 0 Then
            MsgBox("Cannot Proceed. No data in FIN_GAM.  Generate 2059 Files for the branch and run 607", MsgBoxStyle.Information, "Alert!!!")
            Exit Sub
        End If

        Dim SOLID As String
        processmessage("Fetching SOLID")
        sql = "SELECT FIN_SOLID FROM SM WHERE CED_SOLID IN (SELECT LTRIM(BRANCH_NO,'0') FROM MIG_SUMMARY WHERE ROWNUM < 2)"
        Dim cmd2 As New OracleCommand(sql, conn)
        Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
        While dr2.Read
            SOLID = dr2.Item("FIN_SOLID").ToString.Trim
        End While
        dr2.Close()

        oracle_execute_non_query("ten", username, username, "truncate table c_misadv")
        oracle_execute_non_query("ten", username, username, "truncate table c_misdep")
        oracle_execute_non_query("ten", username, username, "truncate table c_misprint")

        processmessage("Generating Package 1")
        sql = "PKGCHECKING.NPAHISTORICALDATA"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
        display_in_File(sql, "C:\du\" & SOLID & "_NPA.txt")

        processmessage("Generating Package 2")
        sql = "PKGCHECKING.NPAHISTORICALDATA1"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.ExecuteNonQuery()

        sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
        display_in_File(sql, "C:\du\" & SOLID & "_NPA1.txt")

        processmessage("Generating Package 3")
        sql = "PKGCHECKING.NPAHISTORICALDATA2"
        Dim cmd6 As New OracleCommand(sql, conn)
        cmd6.CommandType = CommandType.StoredProcedure
        cmd6.ExecuteNonQuery()

        sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
        display_in_File(sql, "C:\du\" & SOLID & "_NPA2.txt")
        cnn.Close()
        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub
    Private Sub option610()   'eNMGB Migration - Create History Transaction Data Dump
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        Dim import_user As String
        Dim uconfirm As String
        Dim dumppath As String

        import_user = InputBox("Username in which data to be imported", "Enter Value", "TEMP")
        dumppath = InputBox("Enter path in which is backup of mig_tran and mig_tran_narration is placed", "Enter Value", "")

        '' drop and create user

        Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\hist_tran_create_user.bat")
        sw0.WriteLine("@echo off")
        sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\static\hist_tran_create_user.sql")
        sw0.Close()
        Process.Start(Disk & ":\dump\script\hist_tran_create_user.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Create Tables

        Dim sw6 As StreamWriter = New StreamWriter(Disk & ":\dump\script\hist_tran_create_tables.bat")
        sw6.WriteLine("@echo off")
        sw6.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\hist_tran_create_tables.sql /nolog ")
        sw6.Close()

        Process.Start(Disk & ":\dump\script\hist_tran_create_tables.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Import dump and moving to base table

        Dim dirs As String() = Directory.GetFiles(dumppath, "*.dmp")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then
            MsgBox("No .dmp exists in the folder " & dumppath, MsgBoxStyle.Critical, "Error")
        Else
            For Each dir In dirs
                tempcount = tempcount + 1

                '' Drop Tables

                Dim sw7 As StreamWriter = New StreamWriter(Disk & ":\dump\script\hist_tran_drop_tables.bat")
                sw7.WriteLine("@echo off")
                sw7.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\hist_tran_drop_tables.sql /nolog ")
                sw7.Close()

                Process.Start(Disk & ":\dump\script\hist_tran_drop_tables.bat")

                uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                If uconfirm <> "Y" Then
                    MsgBox("Exiting application")
                    Exit Sub
                End If

                '' Import Tables

                Dim sw1 As StreamWriter = New StreamWriter(Disk & ":\dump\script\import_user_br.bat")
                sw1.WriteLine("@echo off")
                sw1.WriteLine("imp " & import_user & "/" & import_user & "@ten file=" & dir & "  full=yes")
                sw1.Close()
                Process.Start(Disk & ":\dump\script\import_user_br.bat")

                uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                If uconfirm <> "Y" Then
                    MsgBox("Exiting application")
                    Exit Sub
                End If

                '' Move data

                Dim sw8 As StreamWriter = New StreamWriter(Disk & ":\dump\script\hist_tran_move_data.bat")
                sw8.WriteLine("@echo off")
                sw8.WriteLine("sqlplus " & import_user & "/" & import_user & "@ten @" & Disk & ":\dump\static\hist_tran_move_data.sql /nolog ")
                sw8.Close()

                Process.Start(Disk & ":\dump\script\hist_tran_move_data.bat")

                uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                If uconfirm <> "Y" Then
                    MsgBox("Exiting application")
                    Exit Sub
                End If

            Next
        End If

        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub
    Private Sub option40()   'Export oracle data

        Dim usernames As String
        Dim backuptables As String
        Dim dumpnameprefix As String
        Dim dumppath As String
        Dim uconfirm As String
        Dim uname As String
        Dim sql1 As String

        usernames = InputBox("Enter oracle user name required to backup.  If multiple, seperate by comma", "Enter Value", "")
        backuptables = InputBox("Enter table name required to backup.  If multiple, seperate by comma.  For full backup, enter ALL", "Enter Value", "")
        dumpnameprefix = InputBox("Enter name to suffice after user name in the export backup name", "Enter Value", "")
        dumppath = InputBox("Enter the path in which the dump has to be stored", "Enter Value", "D:\")
        usernames = UCase(usernames)
        backuptables = UCase(backuptables)
        dumpnameprefix = UCase(dumpnameprefix)
        dumppath = UCase(dumppath)

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= cbs;Password= cbs;"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        sql1 = "SELECT UPPER(COLUMN_VALUE) ABCD FROM THE (SELECT CAST(PKGSMGBCOMMON.IN_LIST('" & usernames & "') AS MYTABLETYPE) FROM DUAL)"
        Dim cmd1 As New OracleCommand(sql1, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
        While dr1.Read

            uname = dr1.Item("ABCD").ToString.Trim

            Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\dump\script\export_user_" & uname & ".bat")
            sw2.WriteLine("@echo off")
            If UCase(backuptables) = "ALL" Then
                sw2.WriteLine("exp " & uname & "/" & uname & "@TEN file=" & dumppath & uname & "_" & dumpnameprefix & ".dmp /nolog")
            Else
                sw2.WriteLine("exp " & uname & "/" & uname & "@TEN file=" & dumppath & uname & "_" & dumpnameprefix & ".dmp tables = " & backuptables & " /nolog")
            End If
            sw2.Close()
            Process.Start(Disk & ":\dump\script\export_user_" & uname & ".bat")

            uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
            If uconfirm <> "Y" Then
                MsgBox("Exiting application")
                Exit Sub
            End If

        End While
        dr1.Close()

        oracle_conn.Close()
        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub
    Private Sub option41()   'Backup, Drop and Import Oracle Tables

        Dim tablename As String
        Dim dumpname As String
        Dim uconfirm As String
        Dim uname As String
        Dim sql1 As String
        Dim tname As String

        uname = InputBox("Enter the user name in which the process has to be executed", "Enter Value", "MIG")
        tablename = InputBox("Enter table names involved.  If multiple, seperate by comma", "Enter Value", "CID")
        dumpname = InputBox("Place the dump of table to import in " & Disk & " Drive. Enter the name of the dump without path and extension", "Enter Value", "CID")
        uname = UCase(uname)
        tablename = UCase(tablename)
        dumpname = UCase(dumpname)

        ''Taking backup of the tables to be imported

        Dim sw1 As StreamWriter = New StreamWriter(Disk & ":\dump\script\export_user_" & uname & ".bat")
        sw1.WriteLine("@echo off")
        sw1.WriteLine("exp " & uname & "/" & uname & "@TEN file=" & Disk & ":\" & dumpname & "_" & uname & "_" & "BACKUP.dmp tables = " & tablename & " /nolog")
        sw1.Close()
        Process.Start(Disk & ":\dump\script\export_user_" & uname & ".bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        ''Drop tables

        Dim sw5 As StreamWriter = New StreamWriter(Disk & ":\dump\script\drop_tables.sql")

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= cbs;Password= cbs;"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        sql1 = "SELECT UPPER(COLUMN_VALUE) ABCD FROM THE (SELECT CAST(PKGSMGBCOMMON.IN_LIST('" & tablename & "') AS MYTABLETYPE) FROM DUAL)"
        Dim cmd1 As New OracleCommand(sql1, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
        While dr1.Read

            tname = dr1.Item("ABCD").ToString.Trim
            sw5.WriteLine("drop table " & tname & ";")

        End While
        dr1.Close()

        oracle_conn.Close()

        sw5.Close()

        Dim sw6 As StreamWriter = New StreamWriter(Disk & ":\dump\script\drop_tables.bat")
        sw6.WriteLine("@echo off")
        sw6.WriteLine("sqlplus " & uname & "/" & uname & "@ten @" & Disk & ":\dump\script\drop_tables.sql /nolog ")
        sw6.Close()

        Process.Start(Disk & ":\dump\script\drop_tables.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        '' Importing main data

        Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\dump\script\import.bat")
        sw2.WriteLine("@echo off")
        sw2.WriteLine("imp " & uname & "/" & uname & "@ten file=" & Disk & ":\" & dumpname & ".dmp  full=yes")
        sw2.Close()
        Process.Start(Disk & ":\dump\script\import.bat")
        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Private Sub option42()   'Drop oracle user
        Dim uname As String
        Dim uconfirm As String

        uname = InputBox("Enter the user name to be dropped", "Enter Value", "")

        Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
        sw0.WriteLine("@echo off")
        sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
        sw0.Close()

        Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
        sw.WriteLine("connect sys/sys@ten as sysdba;")
        sw.WriteLine("drop user " & uname & " cascade;")
        sw.WriteLine("create user  " & uname & " identified by  " & uname & " default tablespace domain quota unlimited on domain;")
        sw.WriteLine("grant connect,resource,dba to  " & uname & " ;")
        sw.WriteLine("alter user  " & uname & " temporary tablespace tempdomain;")
        sw.WriteLine("grant create role to  " & uname & " ;")
        sw.WriteLine("grant execute on XMLDOM to  " & uname & " ;")
        sw.WriteLine("grant execute on XMLPARSER to  " & uname & " ;")
        sw.Close()

        Process.Start(Disk & ":\dump\script\create_user.bat")

        uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
        If uconfirm <> "Y" Then
            MsgBox("Exiting application")
            Exit Sub
        End If

        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Private Sub option602() 'eNMGB Migration - Upload Migration Tool Files
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()
        If username.Substring(0, 4).ToUpper <> "MIG0" Then
            MsgBox("Oracle username should start with MIG0")
            Exit Sub
        End If
        If username.Length <> 6 Then
            MsgBox("Invalid Oracle User Name")
            Exit Sub
        End If

        Dim solid As String
        solid = InputBox("Enter the Branch Code")
        solid = solid.PadLeft(5, "0")

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim tempcount As Integer
        processmessage("Checking prerequisites")
        sql = "SELECT COUNT(1) COUNT FROM MIG_SUMMARY"
        Dim cmd1 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd1.ExecuteReader()
        While dr.Read
            tempcount = dr.Item("COUNT").ToString.Trim
        End While
        dr.Close()

        If tempcount = 0 Then
            MsgBox("Cannot Proceed. No data in MIG_Summary.  Run Process 601", MsgBoxStyle.Information, "Alert!!!")
        End If

        Dim flag As Integer = 0

        'Check all files exist
        Dim errmsg As String = " "
        For i = 101 To 115
            Dim checkfilename As String
            'checkfilename = "C:\du\" & solid.Substring(2, 3) & "_816_" & i.ToString() & ".txt"            checkfilename = "C:\du\" & solid.Substring(2, 3) & "_816_" & i.ToString() & ".txt"
            checkfilename = Disk & ":\dump\BR_" & solid & "\MTR_" & solid.Substring(2, 3) & "\" & solid.Substring(2, 3) & "_816_" & i.ToString() & ".txt"
            If File.Exists(checkfilename) = False Then
                errmsg = errmsg & " , " & checkfilename
                flag = 1
            End If
        Next

        If flag = 1 Then
            MsgBox("Following files not found in C:\du folder :" & Environment.NewLine() & errmsg & Environment.NewLine() & Environment.NewLine() & " Place these files and Try again.")
            Exit Sub
        End If

        'Check whether the files are for the same branch
        errmsg = " "
        flag = 0
        For i = 101 To 115
            Dim checkfilename As String
            'checkfilename = "C:\du\" & solid.Substring(2, 3) & "_816_" & i & ".txt"
            checkfilename = Disk & ":\dump\BR_" & solid & "\MTR_" & solid.Substring(2, 3) & "\" & solid.Substring(2, 3) & "_816_" & i.ToString() & ".txt"
            Dim fileReader As System.IO.StreamReader
            fileReader = My.Computer.FileSystem.OpenTextFileReader(checkfilename)
            Dim stringReader As String
            stringReader = fileReader.ReadLine()
            Dim filesolid As String = ""
            Dim startindex As Integer
            Dim endindex As Integer
            startindex = stringReader.IndexOf("|")
            startindex = startindex + 1
            endindex = stringReader.IndexOf("|", startindex)
            filesolid = stringReader.Substring(startindex, endindex - startindex)
            filesolid = filesolid.PadLeft(5, "0")

            If (filesolid <> solid) Then
                errmsg = errmsg & " , " & checkfilename
                flag = 1
            End If
        Next

        errmsg = errmsg.Replace(",", Environment.NewLine())
        If flag = 1 Then
            MsgBox("Mismatch in Filename and File Branch code for the following files : " & Environment.NewLine() & errmsg & Environment.NewLine() & Environment.NewLine() & "Please Correct and try again.")
            Exit Sub
        End If

        oracle_execute_non_query("ten", username, username, "truncate table z_du")

        'Uploading data to MT tables
        For i = 101 To 115
            Dim checkfilename As String
            'checkfilename = "C:\du\" & solid.Substring(2, 3) & "_816_" & i & ".txt"
            checkfilename = Disk & ":\dump\BR_" & solid & "\MTR_" & solid.Substring(2, 3) & "\" & solid.Substring(2, 3) & "_816_" & i.ToString() & ".txt"
            uploadfiledata_without_trim_MigrationTool(checkfilename, username, "N")
        Next

        processmessage("Executing Package")
        sql = "pkgnmgb_mt_data_extraction.process"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.Parameters.Add("BR_CODE", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = solid
        cmd4.ExecuteNonQuery()

        'Check for error in File uploaded file
        flag = 0
        processmessage("Generating Files")
        sql = "SELECT COUNT(1) AS AA FROM C_MISPRINT  WHERE SERIALNO = '000'"
        Dim cmd12 As New OracleCommand(sql, conn)
        Dim dr12 As OracleDataReader = cmd12.ExecuteReader()

        While dr12.Read()
            If dr12("AA").ToString() <> "0" Then
                flag = 1
            End If

        End While
        dr12.Close()

        'Revoking the data entered In case of any issues in data in atlest one file
        If flag = 1 Then
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_BRANCH")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_CITY_CODE1")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_CITY_CODE2")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_CITY_LOCATION")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_CUSTCATEGORY")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_DATAENTRY_SUMM")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_DECEASECODE")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_GUARDIAN")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_LOAN_SANCTION")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_LOANREPAYMENT")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_LOCATION_TABLE")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_LPD")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_NRECODE")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_RELIGION")
            oracle_execute_non_query("ten", username, username, "DELETE FROM MT_STAFFCODE")
            'Generating error file
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO='000'"
            display_in_File(sql, "C:\du\301_Data_Entry_Error.txt")
            Process.Start("C:\du\301_Data_Entry_Error.txt")
        Else
            'Generating data entry status file
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SERIALNO <>'000' ORDER BY SERIALNO,SUBSERIALNO"
            display_in_File(sql, "C:\du\301_Data_Entry_Status.txt")
            Process.Start("C:\du\301_Data_Entry_Status.txt")
        End If
        cnn.Close()
    End Sub

    Private Sub option603() 'eNMGB Migration - Upload CGL File
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        If username.Substring(0, 4).ToUpper <> "MIG0" Then
            MsgBox("Oracle username should start with MIG0")
            Exit Sub
        End If
        If username.Length <> 6 Then
            MsgBox("Invalid Oracle User Name")
            Exit Sub
        End If

        Dim solid As String
        solid = InputBox("Enter the Branch Code")
        solid = solid.PadLeft(5, "0")

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim tempcount As Integer
        processmessage("Checking prerequisites")
        sql = "SELECT COUNT(1) COUNT FROM MT_BRANCH"
        Dim cmd1 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd1.ExecuteReader()
        While dr.Read
            tempcount = dr.Item("COUNT").ToString.Trim
        End While
        dr.Close()

        If tempcount = 0 Then
            MsgBox("Cannot Proceed. No data in mt_branch table.  Run Process ID - 602", MsgBoxStyle.Information, "Alert!!!")
        End If

        If File.Exists(Disk & ":\dump\BR_" & solid & "\CGL.txt") Then
            uploadfiledata_without_trim_MigrationTool(Disk & ":\dump\BR_" & solid & "\CGL.txt", username, "Y")
        Else
            MsgBox("File Named " & Disk & ":\dump\BR_" & solid & "\CGL.txt not found in " & Disk & ":\dump folder", MsgBoxStyle.Information)
            Exit Sub
        End If

        processmessage("Updating CGL Balance")
        sql = "PKGNMGB_CGL_DATA_EXTRACTION.PROCESS"
        Dim cmd2 As New OracleCommand(sql, conn)
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.ExecuteNonQuery()
        cnn.Close()

        sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"

        display_in_File(sql, "C:\du\401_CGL_Data_Entry_Status.txt")
        Process.Start("C:\du\401_CGL_Data_Entry_Status.txt")

    End Sub

    Private Sub option604()   'eNMGB Migration - Assign CustID and Account No
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        Dim tempcount As Integer

        Dim genoption As String = InputBox("<C>ustomer ID" & vbCrLf & "<A>ccount No" & vbCrLf & "<F>ile recreation", "", "")
        If Not (UCase(genoption) = "C" Or UCase(genoption) = "A" Or UCase(genoption) = "F") Then
            MsgBox("Enter valid option. <C> for Customer ID generation; <A> for Account No Generation; <F> Generate File")
            Exit Sub
        End If

        If genoption = "C" Then

            If username.Substring(0, 4).ToUpper <> "MIG0" Then
                MsgBox("Oracle username should start with MIG0")
                Exit Sub
            End If
            If username.Length <> 6 Then
                MsgBox("Invalid Oracle User Name")
                Exit Sub
            End If

            Dim uploadslno As String = InputBox("Update FUF register and enter the latest serial number", "")
            Dim nextpersonlastuploadslno As String = InputBox("Enter upload serial number of next person in queue for CID file", "", "")
            If uploadslno = "" Then
                MsgBox("Enter upload serialno from upload register")
                Exit Sub
            End If
            If nextpersonlastuploadslno = "" Then
                MsgBox("Enter upload serial number of next person in queue for CID file")
                Exit Sub
            End If
            If IsNumeric(uploadslno) = False Then
                MsgBox("Enter proper value in upload serial number field")
                Exit Sub
            End If
            If IsNumeric(nextpersonlastuploadslno) = False Then
                MsgBox("Enter proper value in next person last upload serial number field")
                Exit Sub
            End If

            Dim n_nextpersonlastuploadslno As Integer = nextpersonlastuploadslno
            Dim n_uploadslno As Integer = uploadslno

            If n_nextpersonlastuploadslno >= n_uploadslno Then
                MsgBox("Upload serial number of next person in queue should be less than present upload serial number")
                Exit Sub
            End If


            Dim dirs As String() = Directory.GetFiles("c:\du", "cid.txt")
            Dim dir As String
            Dim totalfiles As Integer
            Dim sql1 As String
            totalfiles = dirs.Length
            If totalfiles = 0 Then
                MsgBox("Place cid.txt in C:/DU and run again", MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If
            For Each dir In dirs

                uploadfiledata(dir, username, "Y")

                processmessage("Updating CID No")
                sql1 = "PKGFUF_FUNCTIONS.UPDATE_CID_FILE_DATA"
                Dim cmd15 As New OracleCommand(sql1, oracle_conn)
                cmd15.CommandType = CommandType.StoredProcedure
                cmd15.Parameters.Add("NEXTSLNO", OracleDbType.Int16, 10, Nothing, ParameterDirection.Input).Value = uploadslno
                cmd15.Parameters.Add("LASTPERSONSLNO", OracleDbType.Int16, 10, Nothing, ParameterDirection.Input).Value = nextpersonlastuploadslno
                cmd15.ExecuteNonQuery()

                sql1 = "SELECT COUNT(1) ABCD FROM C_MISPRINT"
                Dim cmd17 As New OracleCommand(sql1, oracle_conn)
                Dim dr11 As OracleDataReader = cmd17.ExecuteReader()
                While dr11.Read
                    Dim recordcount As String
                    recordcount = dr11(0)
                    If recordcount <> "0" Then
                        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE ROWNUM < 2"
                        Dim cmd8 As New OracleCommand(sql, oracle_conn)
                        Dim dr2 As OracleDataReader = cmd8.ExecuteReader()
                        While dr2.Read
                            Dim errormessage As String
                            errormessage = dr2(0)
                            MsgBox(errormessage)
                            Exit Sub
                        End While
                        dr2.Close()
                    End If
                End While
                dr11.Close()
            Next

            processmessage("Checking whether process 603 has processed...")
            sql = "SELECT COUNT(1) COUNT FROM MIG_CGL"
            Dim cmd1 As New OracleCommand(sql, oracle_conn)
            Dim dr As OracleDataReader = cmd1.ExecuteReader()
            While dr.Read
                tempcount = dr.Item("COUNT").ToString.Trim
            End While
            dr.Close()

            If tempcount = 0 Then
                MsgBox("Cannot Proceed!!! Process 603 is not completed.", MsgBoxStyle.Information, "Alert!!!")
                Exit Sub
            End If

            processmessage("Assigning Cust ID")
            sql = "PKGFUF_FUNCTIONS.ASSIGN_CID_NO"
            Dim cmd3 As New OracleCommand(sql, oracle_conn)
            cmd3.CommandType = CommandType.StoredProcedure
            cmd3.ExecuteNonQuery()

            processmessage("Extracting CID Data")
            sql = "PKGFUF_FUNCTIONS.GENERATE_CID_FILE"
            Dim cmd6 As New OracleCommand(sql, oracle_conn)
            cmd6.CommandType = CommandType.StoredProcedure
            cmd6.Parameters.Add("LASTPERSONSLNO", OracleDbType.Int16, 10, Nothing, ParameterDirection.Input).Value = nextpersonlastuploadslno
            cmd6.ExecuteNonQuery()

            processmessage("Generating CID Data")
            Dim sw1 As StreamWriter = New StreamWriter(folderpath & "\" & "CID.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd7 As New OracleCommand(sql, oracle_conn)
            Dim dr1 As OracleDataReader = cmd7.ExecuteReader()
            While dr1.Read
                Dim linedata As String
                linedata = dr1(0)
                sw1.WriteLine(linedata)
            End While
            dr1.Close()
            sw1.Close()

            MsgBox("Process completed successfully.  Please view the CID file generated for the next person in queue", MsgBoxStyle.Information, "Process Completed")
            Process.Start("C:\du\CID.txt")

        ElseIf genoption = "F" Then

            Dim nextpersonlastuploadslno As String = InputBox("Enter upload serial number of next person in queue", "", "")
            If nextpersonlastuploadslno = "" Then
                MsgBox("Enter upload serial number of next person in queue")
                Exit Sub
            End If
            If IsNumeric(nextpersonlastuploadslno) = False Then
                MsgBox("Enter proper value in next person last upload serial number field")
                Exit Sub
            End If

            sql = "SELECT COUNT(1) ABCD FROM MIG_CUSTID_NO WHERE FIN_CID IS NULL"
            Dim cmd113 As New OracleCommand(sql, oracle_conn)
            Dim dr113 As OracleDataReader = cmd113.ExecuteReader()
            While dr113.Read
                tempcount = dr113.Item("ABCD").ToString.Trim
            End While
            dr113.Close()

            If tempcount <> "0" Then
                MsgBox("Assign Customer ID Number first, using 'C' option", MsgBoxStyle.Information, "Alert!!!")
                Exit Sub
            End If

            processmessage("Extracting CID Data")
            sql = "PKGFUF_FUNCTIONS.GENERATE_CID_FILE"
            Dim cmd6 As New OracleCommand(sql, oracle_conn)
            cmd6.CommandType = CommandType.StoredProcedure
            cmd6.Parameters.Add("LASTPERSONSLNO", OracleDbType.Int16, 10, Nothing, ParameterDirection.Input).Value = nextpersonlastuploadslno
            cmd6.ExecuteNonQuery()

            processmessage("Generating CID Data")
            Dim sw1 As StreamWriter = New StreamWriter(folderpath & "\" & "CID.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd7 As New OracleCommand(sql, oracle_conn)
            Dim dr1 As OracleDataReader = cmd7.ExecuteReader()
            While dr1.Read
                Dim linedata As String
                linedata = dr1(0)
                sw1.WriteLine(linedata)
            End While
            dr1.Close()
            sw1.Close()

            sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
            Dim cmd123 As New OracleCommand(sql, oracle_conn)
            Dim dr123 As OracleDataReader = cmd123.ExecuteReader()
            While dr123.Read
                tempcount = dr123.Item("ABCD").ToString.Trim
            End While
            dr123.Close()

            MsgBox("Process completed successfully.  Please view the CID file recreated", MsgBoxStyle.Information, "Process Completed")
            Process.Start("C:\du\CID.txt")

        ElseIf genoption = "A" Then

            sql = "SELECT COUNT(1) ABCD FROM MIG_CUSTID_NO WHERE FIN_CID IS NULL"
            Dim cmd113 As New OracleCommand(sql, oracle_conn)
            Dim dr113 As OracleDataReader = cmd113.ExecuteReader()
            While dr113.Read
                tempcount = dr113.Item("ABCD").ToString.Trim
            End While
            dr113.Close()

            If tempcount <> "0" Then
                MsgBox("Assign Customer ID Number first, using 'C' option", MsgBoxStyle.Information, "Alert!!!")
                Exit Sub
            End If

            processmessage("Assigning Account No")
            sql = "PKGFUF_FUNCTIONS.ASSIGN_AC_NO"
            Dim cmd4 As New OracleCommand(sql, oracle_conn)
            cmd4.CommandType = CommandType.StoredProcedure
            cmd4.ExecuteNonQuery()

            processmessage("Extracting TBA Data")
            sql = "PKGFUF_FUNCTIONS.EXTRACT_TBA_DATA"
            Dim cmd5 As New OracleCommand(sql, oracle_conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.ExecuteNonQuery()

            MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

        End If

        oracle_conn.Close()

    End Sub
    Private Sub option605()   'eNMGB Migration - FUF Generation
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        If username.Substring(0, 4).ToUpper <> "MIG0" Then
            MsgBox("Oracle username should start with MIG0")
            Exit Sub
        End If
        If username.Length <> 6 Then
            MsgBox("Invalid Oracle User Name")
            Exit Sub
        End If

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        Dim branchcode As String
        Dim branchname As String
        Dim cedgebranchcode As String

        '' Fetching branch codes

        Dim sql1 As String
        sql1 = "SELECT PKGFUF_FUNCTIONS.FINACLE_SOLID(BRANCH_NO) ABCD, BRANCH_NO FROM MIG_SUMMARY WHERE ROWNUM < 2"
        Dim cmd1 As New OracleCommand(sql1, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
        While dr1.Read
            branchcode = dr1.Item("ABCD").ToString.Trim
            cedgebranchcode = dr1.Item("BRANCH_NO").ToString.Trim
        End While
        dr1.Close()

        '' Fetching branch name

        Dim sql10 As String
        sql10 = "SELECT CED_BRANCHNAME FROM SM WHERE FIN_SOLID = '" & branchcode & "'"
        Dim cmd10 As New OracleCommand(sql10, oracle_conn)
        Dim dr10 As OracleDataReader = cmd10.ExecuteReader()
        While dr10.Read
            branchname = dr10.Item("CED_BRANCHNAME").ToString.Trim
        End While
        dr10.Close()

        '' Fetching data in mt_zenith table

        Dim tempcount As Integer
        processmessage("Checking prerequisites")
        Dim sql13 As String
        sql13 = "SELECT COUNT(1) COUNT FROM MT_ZENITH"
        Dim cmd13 As New OracleCommand(sql13, oracle_conn)
        Dim dr13 As OracleDataReader = cmd13.ExecuteReader()
        While dr13.Read
            tempcount = dr13.Item("COUNT").ToString.Trim
        End While
        dr13.Close()

        Dim tempvarchar As String
        processmessage("Checking prerequisites")
        Dim sql23 As String
        sql23 = "SELECT PKGFUF_FUNCTIONS.NMGB_TBA_BRANCH_FLAG( " & cedgebranchcode & ") COUNT FROM DUAL"
        Dim cmd23 As New OracleCommand(sql23, oracle_conn)
        Dim dr23 As OracleDataReader = cmd23.ExecuteReader()
        While dr23.Read
            tempvarchar = dr23.Item("COUNT").ToString.Trim
        End While
        dr23.Close()


        If tempvarchar = "Y" And tempcount = 0 Then
            MsgBox("Cannot Proceed. No data in MT_Zenith table.  Run Process ID - 604", MsgBoxStyle.Information, "Alert!!!")
            Exit Sub
        End If

        '' Fetching whether account number is assigned

        Dim tempcount1 As Integer
        processmessage("Checking prerequisites")
        Dim sql14 As String
        sql14 = "SELECT COUNT(1) COUNT FROM MIG_ACC_NO WHERE FORACID IS NOT NULL"
        Dim cmd14 As New OracleCommand(sql14, oracle_conn)
        Dim dr14 As OracleDataReader = cmd14.ExecuteReader()
        While dr14.Read
            tempcount1 = dr14.Item("COUNT").ToString.Trim
        End While
        dr14.Close()

        If tempcount1 = 0 Then
            MsgBox("Cannot Proceed. Account number not assigned.  Run Process ID - 604", MsgBoxStyle.Information, "Alert!!!")
            Exit Sub
        End If

        '' Fetching whether customer id is assigned

        Dim tempcount2 As Integer
        processmessage("Checking prerequisites")
        Dim sql15 As String
        sql15 = "SELECT COUNT(1) COUNT FROM MIG_CUSTID_NO WHERE FIN_CID IS NOT NULL"
        Dim cmd15 As New OracleCommand(sql15, oracle_conn)
        Dim dr15 As OracleDataReader = cmd14.ExecuteReader()
        While dr15.Read
            tempcount2 = dr15.Item("COUNT").ToString.Trim
        End While
        dr15.Close()

        If tempcount2 = 0 Then
            MsgBox("Cannot Proceed. Customer ID not assigned.  Run Process ID - 604", MsgBoxStyle.Information, "Alert!!!")
            Exit Sub
        End If

        '' Accepting values from user

        Dim reportid As String = InputBox("Enter Report ID", "Enter Value", "ALL")
        Dim migdate As Date = InputBox("Enter Date of Migration (DD/MM/YYYY)", "Enter Value", "01-11-2014")
        ''30-06-2014
        Dim fufname As String
        Dim procedure_len As Integer
        procedure_len = txtmenu.Text.Length

        Dim folderpath2 As String = "c:\du\" & branchcode & " " & branchname & "\" & branchcode & " " & "FUF Generation"

        Dim sql6 As String
        Dim startid As String
        Dim endid As String

        If IsNumeric(reportid.Substring(0, 1)) = True Then

            startid = reportid.Substring(0, 2)
            endid = reportid.Substring(2, 2)
            sql6 = "SELECT SLNO,REPORT_ID FROM FI WHERE REPORT_TYPE = 'FUF' AND SLNO BETWEEN " & startid & " AND " & endid & " ORDER BY SLNO"

        ElseIf reportid = "ALL" Then

            sql6 = "SELECT SLNO,REPORT_ID,REPORT_DESC,REPORT_FILE_NAME FROM FI WHERE REPORT_TYPE = 'FUF' ORDER BY SLNO"

        Else

            sql6 = "SELECT SLNO,REPORT_ID  FROM FI WHERE REPORT_TYPE = 'FUF' AND REPORT_ID =  '" & reportid & "'"

        End If

        Dim cmd6 As New OracleCommand(sql6, oracle_conn)
        Dim dr6 As OracleDataReader = cmd6.ExecuteReader()
        While dr6.Read

            Dim reportid1 As String
            Dim fileno As String
            reportid1 = dr6.Item("REPORT_ID").ToString.Trim
            fileno = dr6.Item("SLNO").ToString.Trim

            If Not Directory.Exists(folderpath2) Then
                Directory.CreateDirectory(folderpath2)
            End If

        End While
        dr6.Close()



        Dim sql2 As String
        If IsNumeric(reportid.Substring(0, 1)) = True Then

            startid = reportid.Substring(0, 2)
            endid = reportid.Substring(2, 2)
            sql2 = "SELECT SLNO,REPORT_ID,REPORT_DESC,REPORT_FILE_NAME FROM FI WHERE REPORT_TYPE = 'FUF' AND SLNO BETWEEN " & startid & " AND " & endid & " ORDER BY SLNO"

        ElseIf reportid = "ALL" Then

            sql2 = "SELECT SLNO,REPORT_ID,REPORT_DESC,REPORT_FILE_NAME FROM FI WHERE REPORT_TYPE = 'FUF' ORDER BY SLNO"

        Else

            sql2 = "SELECT SLNO,REPORT_ID,REPORT_DESC,REPORT_FILE_NAME FROM FI WHERE REPORT_TYPE = 'FUF' AND REPORT_ID =  '" & reportid & "'"

        End If

        If IsNumeric(reportid.Substring(0, 1)) = True Then

            Dim cmd2 As New OracleCommand(sql2, oracle_conn)
            Dim dr2 As OracleDataReader = cmd2.ExecuteReader()

            oracle_execute_non_query("ten", username, username, "TRUNCATE TABLE C_MISPRINT")

            While dr2.Read

                Dim reportid1 As String
                Dim reportdesc As String
                Dim fileno As String
                Dim uconfirm As String
                reportid1 = dr2.Item("REPORT_ID").ToString.Trim
                reportdesc = dr2.Item("REPORT_DESC").ToString.Trim
                fileno = dr2.Item("SLNO").ToString.Trim

                If reportid1 = "EMI" Then


                    '' Rebuilding Index 

                    Dim sw7 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_table4.bat")
                    sw7.WriteLine("@echo off")
                    sw7.WriteLine("sqlplus " & username & "/" & username & "@ten @" & Disk & ":\dump\static\Table4.sql /nolog ")
                    sw7.Close()
                    Process.Start(Disk & ":\dump\script\create_table4.bat")
                    uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                    If uconfirm <> "Y" Then
                        MsgBox("Exiting application")
                        Exit Sub
                    End If

                    processmessage("Calculating EMI")
                    CalculateEMI()

                    '' Creating index in fuf_lrs_1

                    Dim sw17 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_table3.bat")
                    sw17.WriteLine("@echo off")
                    sw17.WriteLine("sqlplus " & username & "/" & username & "@ten @" & Disk & ":\dump\static\Table3.sql /nolog ")
                    sw17.Close()
                    Process.Start(Disk & ":\dump\script\create_table3.bat")
                    uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                    If uconfirm <> "Y" Then
                        MsgBox("Exiting application")
                        Exit Sub
                    End If

                Else

                    processmessage("Processing File No: " & fileno & " - " & reportid1 & " - " & reportdesc)
                    sql = "PKGFUF.GENERATE_FUF"
                    Dim cmd3 As New OracleCommand(sql, oracle_conn)
                    cmd3.CommandType = CommandType.StoredProcedure
                    cmd3.Parameters.Add("FUFNAME", OracleDbType.Varchar2, 50, Nothing, ParameterDirection.Input).Value = Trim(reportid1)
                    cmd3.Parameters.Add("MDATE", OracleDbType.Date, 50, Nothing, ParameterDirection.Input).Value = migdate
                    cmd3.ExecuteNonQuery()

                End If

            End While
            dr2.Close()

        ElseIf reportid = "ALL" Then

            Dim cmd2 As New OracleCommand(sql2, oracle_conn)
            Dim dr2 As OracleDataReader = cmd2.ExecuteReader()

            oracle_execute_non_query("ten", username, username, "TRUNCATE TABLE C_MISPRINT")

            While dr2.Read

                Dim reportid1 As String
                Dim reportdesc As String
                Dim fileno As String
                Dim uconfirm As String
                reportid1 = dr2.Item("REPORT_ID").ToString.Trim
                reportdesc = dr2.Item("REPORT_DESC").ToString.Trim
                fileno = dr2.Item("SLNO").ToString.Trim

                If reportid1 = "EMI" Then

                    '' Rebuilding Index 

                    Dim sw7 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_table4.bat")
                    sw7.WriteLine("@echo off")
                    sw7.WriteLine("sqlplus " & username & "/" & username & "@ten @" & Disk & ":\dump\static\Table4.sql /nolog ")
                    sw7.Close()
                    Process.Start(Disk & ":\dump\script\create_table4.bat")
                    uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                    If uconfirm <> "Y" Then
                        MsgBox("Exiting application")
                        Exit Sub
                    End If

                    processmessage("Calculating EMI")
                    CalculateEMI()

                    '' Creating index in fuf_lrs_1

                    Dim sw17 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_table3.bat")
                    sw17.WriteLine("@echo off")
                    sw17.WriteLine("sqlplus " & username & "/" & username & "@ten @" & Disk & ":\dump\static\Table3.sql /nolog ")
                    sw17.Close()
                    Process.Start(Disk & ":\dump\script\create_table3.bat")
                    uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
                    If uconfirm <> "Y" Then
                        MsgBox("Exiting application")
                        Exit Sub
                    End If

                Else

                    processmessage("Processing File No: " & fileno & " - " & reportid1 & " - " & reportdesc)
                    sql = "PKGFUF.GENERATE_FUF"
                    Dim cmd3 As New OracleCommand(sql, oracle_conn)
                    cmd3.CommandType = CommandType.StoredProcedure
                    cmd3.Parameters.Add("FUFNAME", OracleDbType.Varchar2, 50, Nothing, ParameterDirection.Input).Value = Trim(reportid1)
                    cmd3.Parameters.Add("MDATE", OracleDbType.Date, 50, Nothing, ParameterDirection.Input).Value = migdate
                    cmd3.ExecuteNonQuery()

                End If

            End While
            dr2.Close()

        Else

            Dim cmd2 As New OracleCommand(sql2, oracle_conn)
            Dim dr2 As OracleDataReader = cmd2.ExecuteReader()

            While dr2.Read()
                Dim reportid1 As String
                Dim reportdesc As String
                Dim fileno As String
                reportid1 = dr2.Item("REPORT_ID").ToString.Trim
                reportdesc = dr2.Item("REPORT_DESC").ToString.Trim
                fileno = dr2.Item("SLNO").ToString.Trim

                oracle_execute_non_query("ten", username, username, "TRUNCATE TABLE C_MISPRINT")

                processmessage("Processing File No: " & fileno & " - " & reportid1 & " - " & reportdesc)
                sql = "PKGFUF.GENERATE_FUF"
                Dim cmd3 As New OracleCommand(sql, oracle_conn)
                cmd3.CommandType = CommandType.StoredProcedure
                cmd3.Parameters.Add("FUFNAME", OracleDbType.Varchar2, 50, Nothing, ParameterDirection.Input).Value = Trim(reportid1)
                cmd3.Parameters.Add("MDATE", OracleDbType.Date, 50, Nothing, ParameterDirection.Input).Value = migdate
                cmd3.ExecuteNonQuery()

            End While
            dr2.Close()

        End If

        '' CREATIING FILES IN "Wipro" FOLDER
        Dim sql8 As String
        If IsNumeric(reportid.Substring(0, 1)) = True Then

            startid = reportid.Substring(0, 2)
            endid = reportid.Substring(2, 2)
            sql8 = "SELECT SLNO,REPORT_ID FROM FI WHERE REPORT_TYPE = 'FUF' AND SLNO BETWEEN " & startid & " AND " & endid & " ORDER BY REPORT_ID"

        ElseIf reportid = "ALL" Then

            sql8 = "SELECT SLNO,REPORT_ID FROM FI WHERE REPORT_TYPE = 'FUF' ORDER BY REPORT_ID"

        Else

            sql8 = "SELECT SLNO,REPORT_ID FROM FI WHERE REPORT_TYPE = 'FUF' AND REPORT_ID =  '" & reportid & "'"

        End If


        Dim sw3 As StreamWriter
        sw3 = New StreamWriter(folderpath2 & "\" & branchcode & "COUNT.LST")
        sw3.WriteLine("--------------------------------------------------------")
        sw3.WriteLine("FUF Name                                          Count")
        sw3.WriteLine("--------------------------------------------------------")
        Dim cmd8 As New OracleCommand(sql8, oracle_conn)
        Dim dr8 As OracleDataReader = cmd8.ExecuteReader()
        While dr8.Read

            fufname = dr8.Item("REPORT_ID").ToString.Trim
            Dim nooffiles As Integer
            If fufname = "AC1" Or fufname = "SBB" Or fufname = "SBI" Then

                Dim sql101 As String
                sql101 = "SELECT ROUND((COUNT(1)/8000)+.50) ABCD FROM C_MISPRINT  WHERE SOLID = '" & fufname & "'"
                Dim cmd101 As New OracleCommand(sql101, oracle_conn)
                Dim dr101 As OracleDataReader = cmd101.ExecuteReader()
                tempcount = 0
                While dr101.Read
                    nooffiles = dr101(0)
                End While
                dr101.Close()

                Dim startno As Integer
                Dim endno As Integer
                Dim kkkk As String
                Dim aaaa As Integer

                For i As Integer = 1 To nooffiles
                    'oracle_fields(i) = dr(i)
                    Dim sw1 As StreamWriter
                    sw1 = New StreamWriter(folderpath2 & "\" & branchcode & fufname & "_" & i & ".LST")
                    startno = 1 + (8000 * (i - 1))
                    endno = 8000 * i
                    Dim sql66 As String
                    sql66 = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SOLID = '" & fufname & "' ORDER BY SERIALNO,SUBSERIALNO"
                    Dim cmd66 As New OracleCommand(sql66, oracle_conn)
                    Dim dr66 As OracleDataReader = cmd66.ExecuteReader()
                    tempcount = 0
                    processmessage("Generating File " & branchcode & fufname & "_" & i & ".LST")
                    aaaa = 0
                    While dr66.Read
                        Dim linedata3 As String
                        tempcount = tempcount + 1
                        If tempcount >= startno And tempcount <= endno Then
                            aaaa = aaaa + 1
                            linedata3 = dr66(0)
                            sw1.WriteLine(linedata3)
                        End If
                    End While
                    kkkk = fufname & "_" & i
                    sw3.WriteLine(kkkk.PadRight(50) & aaaa)
                    dr66.Close()
                    sw1.Close()
                Next
            Else

                Dim sw1 As StreamWriter
                sw1 = New StreamWriter(folderpath2 & "\" & branchcode & fufname & ".LST")
                Dim sql66 As String
                sql66 = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SOLID = '" & fufname & "' ORDER BY SERIALNO,SUBSERIALNO"
                Dim cmd66 As New OracleCommand(sql66, oracle_conn)
                Dim dr66 As OracleDataReader = cmd66.ExecuteReader()
                tempcount = 0
                While dr66.Read
                    Dim linedata3 As String
                    tempcount = tempcount + 1
                    linedata3 = dr66(0)
                    processmessage("Generating File " & branchcode & fufname & ".LST")
                    sw1.WriteLine(linedata3)
                End While
                sw3.WriteLine(fufname.PadRight(50) & tempcount)
                dr66.Close()
                sw1.Close()

            End If

        End While
        dr8.Close()
        sw3.WriteLine("--------------------------------------------------------")
        sw3.Close()

        oracle_conn.Close()
        MsgBox("Files generated Successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Private Sub option606()   'eNMGB Migration - Data Check Reports
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        Dim mdate As String = InputBox("Enter migration date", "", "01-11-2014")
        ''30-06-2014
        Try

            RptDate = CDate(mdate)

        Catch ex As Exception

            MsgBox("Enter valid date", MsgBoxStyle.Critical, "Invalid date")
            Exit Sub

        End Try
        If username.Substring(0, 4).ToUpper <> "MIG0" Then
            MsgBox("Oracle username should start with MIG0")
            Exit Sub
        End If
        If username.Length <> 6 Then
            MsgBox("Invalid Oracle User Name")
            Exit Sub
        End If
        Dim reportcategory As String = InputBox("Enter Report Category", "Enter Value", "ALL")
        Dim reportid As String = InputBox("Enter Report ID", "Enter Value", "ALL")

        Dim oracle_cnn_string As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        Dim tempcount As Integer
        processmessage("Checking prerequisites")
        Dim sql1 As String
        sql1 = "SELECT COUNT(1) COUNT FROM FUF_AC1"
        Dim cmd1 As New OracleCommand(sql1, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
        While dr1.Read
            tempcount = dr1.Item("COUNT").ToString.Trim
        End While
        dr1.Close()
        If tempcount = 0 Then
            MsgBox("Cannot Proceed!!! Execute Process ID 605", MsgBoxStyle.Information, "Alert!!!")
            Exit Sub
        Else
            processmessage("")
        End If

        processmessage("Retrieving Finacle SOLID")
        Dim sql2 As String
        sql2 = "SELECT PKGFUF_FUNCTIONS.FINACLE_SOLID(BRANCH_NO) BRANCH_NO FROM MIG_SUMMARY WHERE ROWNUM < 2"
        Dim cmd2 As New OracleCommand(sql2, oracle_conn)
        Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
        Dim branchcode As String
        While dr2.Read
            branchcode = dr2.Item("BRANCH_NO").ToString.Trim
        End While
        dr2.Close()

        processmessage("Retrieving CEDGE Branch Name")
        Dim sql3 As String
        sql3 = "SELECT CED_BRANCHNAME FROM SM WHERE FIN_SOLID = '" & branchcode & "'"
        Dim cmd3 As New OracleCommand(sql3, oracle_conn)
        Dim dr3 As OracleDataReader = cmd3.ExecuteReader()
        Dim branchname As String
        While dr3.Read
            branchname = dr3.Item("CED_BRANCHNAME").ToString.Trim
        End While
        dr3.Close()

        processmessage("Retrieving CEDGE Branch Code")
        Dim sql4 As String
        sql4 = "SELECT BRANCH_NO FROM MIG_SUMMARY WHERE ROWNUM < 2"
        Dim cmd4 As New OracleCommand(sql4, oracle_conn)
        Dim dr4 As OracleDataReader = cmd4.ExecuteReader()
        Dim cedgebrcode As String
        While dr4.Read
            cedgebrcode = dr4.Item("BRANCH_NO").ToString.Trim
        End While
        dr4.Close()

        Dim folderpath As String
        Dim filename As String
        Dim ReportHead As String
        Dim ReportHeadDesc As String
        Dim repid As String
        Dim repdesc As String
        Dim repfilename As String

        Dim sql5 As String
        Dim sql6 As String
        Dim sql7 As String
        Dim sql8 As String

        If reportcategory = "ALL" Then
            sql5 = "SELECT DISTINCT REPORT_TYPE,REPORT_DESC FROM RI WHERE REPORT_ID = '0'"
        Else
            sql5 = "SELECT DISTINCT REPORT_TYPE,REPORT_DESC FROM RI WHERE REPORT_TYPE = '" & reportcategory & "' AND REPORT_ID = '0'"
        End If
        Dim cmd5 As New OracleCommand(sql5, oracle_conn)
        Dim dr5 As OracleDataReader = cmd5.ExecuteReader()
        While dr5.Read
            ReportHead = dr5.Item("REPORT_TYPE").ToString.Trim
            ReportHeadDesc = dr5.Item("REPORT_DESC").ToString.Trim
            folderpath = "c:\du\" & branchcode & " " & branchname & "\" & branchcode & " " & ReportHeadDesc
            If Not Directory.Exists(folderpath) Then
                Directory.CreateDirectory(folderpath)
            End If
            If reportid = "ALL" Then
                sql6 = "SELECT REPORT_ID,REPORT_DESC,REPORT_FILE_NAME FROM RI WHERE REPORT_TYPE = '" & ReportHead & "' AND REPORT_ID <> '0' ORDER BY REPORT_ID"
            Else
                sql6 = "SELECT REPORT_ID,REPORT_DESC,REPORT_FILE_NAME FROM RI WHERE REPORT_TYPE = '" & ReportHead & "' AND REPORT_ID = '" & reportid & "'"
            End If
            Dim cmd6 As New OracleCommand(sql6, oracle_conn)
            Dim dr6 As OracleDataReader = cmd6.ExecuteReader()
            While dr6.Read
                repid = dr6.Item("REPORT_ID").ToString.Trim
                repdesc = dr6.Item("REPORT_DESC").ToString.Trim
                repfilename = dr6.Item("REPORT_FILE_NAME").ToString.Trim
                filename = repid & " " & repfilename & ".txt"
                If File.Exists(folderpath & "\" & filename) Then
                    File.Delete(folderpath & "\" & filename)
                End If
                processmessage("Processing Report : " & ReportHead & " " & repid & " " & repdesc)
                sql7 = "PKGCHECKING.DATA_CHECK"
                Dim cmd7 As New OracleCommand(sql7, oracle_conn)
                cmd7.CommandType = CommandType.StoredProcedure
                cmd7.Parameters.Add("REPORTID", OracleDbType.Varchar2, 50, Nothing, ParameterDirection.Input).Value = repid
                cmd7.Parameters.Add("MDATE", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
                cmd7.ExecuteNonQuery()

                processmessage("Processing Report : " & ReportHead & " " & repid & " " & repdesc & " - Generating Report")
                Dim sw8 As StreamWriter = New StreamWriter(folderpath & "\" & filename)
                sql8 = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
                Dim cmd8 As New OracleCommand(sql8, oracle_conn)
                Dim dr8 As OracleDataReader = cmd8.ExecuteReader()
                While dr8.Read
                    Dim linedata As String
                    linedata = dr8(0)
                    sw8.WriteLine(linedata)
                End While
                dr8.Close()
                sw8.Close()
            End While
            dr6.Close()
        End While
        dr5.Close()
        oracle_conn.Close()
        processmessage("")
        MsgBox("Files generated Successfully", MsgBoxStyle.Information, "Process Completed")
    End Sub

    Private Sub option39()   'Migration Tool Data Entry Status Email

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:\du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

        End If

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        'Calling packages

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - Data ID - 1132")      'Migration Tool Data Entry Status

        sql = "PKGEMAIL113.DATAID_1132"
        Dim cmd7 As New OracleCommand(sql, conn)
        cmd7.CommandType = CommandType.StoredProcedure
        cmd7.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "BR"
        cmd7.ExecuteNonQuery()

        sendemail("kgbmis1@gmail.com", "ten", username, username)
        'sendemail("dipsdot@gmail.com", "ten", username, username)

    End Sub

    Public Sub mirrorimage_source_destination(ByVal source, ByVal destination)
        Dim sourcefolder As String = ""
        Dim destinationfolder As String = ""
        Dim sourcepath() As String

        If source.Contains("\") Then
            sourcepath = source.Split("\")
            sourcefolder = sourcepath(sourcepath.Length - 1)
        End If

        If destination.Contains("\") Then
            sourcepath = destination.Split("\")
            destinationfolder = sourcepath(sourcepath.Length - 1)
        End If

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim sql As String
        sql = "SELECT SOLID,REPORTDATA FROM C_MISPRINT"
        Dim cmd4 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd4.ExecuteReader()
        If dr.Read() = True Then
            'Source to Destination Copying
            sql = "SELECT SOLID,REPORTDATA FROM C_MISPRINT WHERE TRIM(SOLID) ='DS'"
            Dim cmd5 As New OracleCommand(sql, conn)
            Dim dr1 As OracleDataReader = cmd5.ExecuteReader()

            While dr1.Read()
                Dim fullpath As String
                fullpath = ""
                fullpath = dr1(1).ToString()
                If fullpath.Contains(sourcefolder) Then
                    fullpath = fullpath.Remove(0, fullpath.IndexOf(sourcefolder) + Len(sourcefolder) + 1)
                End If
                Dim items_inpath As String()
                items_inpath = fullpath.Split("\")
                Dim createdpath As String
                createdpath = ""
                For i As Integer = 0 To items_inpath.Length - 1
                    If createdpath <> "" Then
                        If Directory.Exists(destination & "\" & createdpath & "\" & items_inpath(i)) = False Then
                            Directory.CreateDirectory(destination & "\" & createdpath & "\" & items_inpath(i))
                        End If
                        createdpath = createdpath & "\" & items_inpath(i)
                    Else
                        If Directory.Exists(destination & "\" & items_inpath(i)) = False Then
                            Directory.CreateDirectory(destination & "\" & items_inpath(i))
                        End If
                        createdpath = items_inpath(i)
                    End If

                Next

            End While
            dr1.Close()

            sql = "SELECT SOLID,REPORTDATA FROM C_MISPRINT WHERE TRIM(SOLID) ='FS'"
            Dim cmd6 As New OracleCommand(sql, conn)
            Dim dr2 As OracleDataReader = cmd6.ExecuteReader()

            While dr2.Read()
                Dim fullpath As String
                fullpath = ""
                fullpath = dr2(1).ToString()
                If fullpath.Contains(sourcefolder) Then
                    fullpath = fullpath.Remove(0, fullpath.IndexOf(sourcefolder) + Len(sourcefolder) + 1)
                End If
                Dim items_inpath As String()
                items_inpath = fullpath.Split("\")
                Dim createdpath As String
                createdpath = ""
                For i As Integer = 0 To items_inpath.Length - 2
                    If createdpath <> "" Then
                        If Directory.Exists(destination & "\" & createdpath & "\" & items_inpath(i)) = False Then
                            Directory.CreateDirectory(destination & "\" & createdpath & "\" & items_inpath(i))
                        End If
                        createdpath = createdpath & "\" & items_inpath(i)
                    Else
                        If Directory.Exists(destination & "\" & items_inpath(i)) = False Then
                            Directory.CreateDirectory(destination & "\" & items_inpath(i))
                        End If
                        createdpath = items_inpath(i)
                    End If
                Next

                If File.Exists(destination & "\" & createdpath & "\" & items_inpath(items_inpath.Length - 1)) Then
                    File.Delete(destination & "\" & createdpath & "\" & items_inpath(items_inpath.Length - 1))
                End If
                File.Copy(dr2(1).ToString, destination & "\" & createdpath & "\" & items_inpath(items_inpath.Length - 1))


            End While
            dr2.Close()
        End If
        dr.Close()
    End Sub
    Public Sub createbackupofnewfiles(ByVal destination, ByVal backup_folder)
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim sql As String
        Dim directoryname As String
        sql = "SELECT SOLID,REPORTDATA FROM C_MISPRINT"
        Dim cmd4 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd4.ExecuteReader()
        If dr.Read() = True Then
            If Directory.Exists(destination) Then

                directoryname = System.DateTime.Now
                directoryname = Format(System.DateTime.Now, "yyyy-MM-dd HH:mm:ss")
                directoryname = directoryname.Replace("-", "")
                directoryname = directoryname.Replace("\", "")
                directoryname = directoryname.Replace(":", "")
                directoryname = directoryname.Replace(" ", "")
                Directory.CreateDirectory(destination & "\" & directoryname)

                sql = "SELECT SOLID,REPORTDATA FROM C_MISPRINT WHERE TRIM(SOLID) ='D'"
                Dim cmd5 As New OracleCommand(sql, conn)
                Dim dr1 As OracleDataReader = cmd5.ExecuteReader()

                While dr1.Read()
                    Dim fullpath As String
                    fullpath = ""
                    fullpath = dr1(1).ToString()
                    If fullpath.Contains(backup_folder) Then
                        fullpath = fullpath.Remove(0, fullpath.IndexOf(backup_folder))
                    End If
                    Dim items_inpath As String()
                    items_inpath = fullpath.Split("\")
                    Dim createdpath As String
                    createdpath = ""
                    For i As Integer = 0 To items_inpath.Length - 1
                        If createdpath <> "" Then
                            If Directory.Exists(destination & "\" & directoryname & "\" & createdpath & "\" & items_inpath(i)) = False Then
                                Directory.CreateDirectory(destination & "\" & directoryname & "\" & createdpath & "\" & items_inpath(i))
                            End If
                            createdpath = createdpath & "\" & items_inpath(i)
                        Else
                            If Directory.Exists(destination & "\" & directoryname & "\" & items_inpath(i)) = False Then
                                Directory.CreateDirectory(destination & "\" & directoryname & "\" & items_inpath(i))
                            End If
                            createdpath = items_inpath(i)
                        End If

                    Next

                End While
            End If


            sql = "SELECT SOLID,REPORTDATA FROM C_MISPRINT WHERE TRIM(SOLID) ='F'"
            Dim cmd6 As New OracleCommand(sql, conn)
            Dim dr2 As OracleDataReader = cmd6.ExecuteReader()

            While dr2.Read()
                Dim fullpath As String
                fullpath = ""
                fullpath = dr2(1).ToString()
                If fullpath.Contains(backup_folder) Then
                    fullpath = fullpath.Remove(0, fullpath.IndexOf(backup_folder))
                End If
                Dim items_inpath As String()
                items_inpath = fullpath.Split("\")
                Dim createdpath As String
                createdpath = ""
                For i As Integer = 0 To items_inpath.Length - 2
                    If createdpath <> "" Then
                        If Directory.Exists(destination & "\" & directoryname & "\" & createdpath & "\" & items_inpath(i)) = False Then
                            Directory.CreateDirectory(destination & "\" & directoryname & "\" & createdpath & "\" & items_inpath(i))
                        End If
                        createdpath = createdpath & "\" & items_inpath(i)
                    Else
                        If Directory.Exists(destination & "\" & directoryname & "\" & items_inpath(i)) = False Then
                            Directory.CreateDirectory(destination & "\" & directoryname & "\" & items_inpath(i))
                        End If
                        createdpath = items_inpath(i)
                    End If
                Next

                File.Copy(dr2(1).ToString, destination & "\" & directoryname & "\" & createdpath & "\" & items_inpath(items_inpath.Length - 1))

            End While
        End If
        dr.Close()
    End Sub
    Public Sub directory_listing(ByVal type, ByVal folderpath, ByVal backup_folder, ByVal caller)
        Dim type_folder As String
        Dim timestamp As String
        Dim size As Integer
        Dim tabname As String
        Dim comparepart As String

        If type = "Source" Then
            tabname = "C_MISADV"
        Else
            tabname = "C_MISDEP"
        End If

        For Each directoryname As String In Directory.GetDirectories(folderpath)
            type_folder = "D"
            timestamp = Directory.GetCreationTime(directoryname).ToString()
            size = 0
            comparepart = ""
            If caller = "Differential" Then
                If directoryname.Contains(backup_folder) Then
                    comparepart = directoryname.Remove(0, directoryname.IndexOf(backup_folder))
                End If
            End If

            oracle_execute_non_query("ten", username, username, "INSERT INTO " & tabname & "(MEMO1,TEXT2,TEXT3,TEXT4,NUMBER1) values(' " & directoryname & "','" & type_folder & "','" & timestamp & "','" & comparepart & "'," & size & ")")
            If type = "Source" Then
                directory_listing("Source", directoryname, backup_folder, caller)
            Else
                directory_listing("Destination", directoryname, backup_folder, caller)
            End If
        Next

        For Each directoryname As String In Directory.GetFiles(folderpath)
            type_folder = "F"
            timestamp = File.GetLastWriteTime(directoryname).ToString()
            Dim file1 As New FileInfo(directoryname)
            size = file1.Length
            comparepart = ""
            If caller = "Differential" Then
                If directoryname.Contains(backup_folder) Then
                    comparepart = directoryname.Remove(0, directoryname.IndexOf(backup_folder))
                End If
            End If

            oracle_execute_non_query("ten", username, username, "INSERT INTO  " & tabname & " (MEMO1,TEXT2,TEXT3,TEXT4,NUMBER1) values(' " & directoryname & "','" & type_folder & "','" & timestamp & "','" & comparepart & "'," & size & ")")
        Next

    End Sub
    Public Sub option827()   'Upload files extension based
        Dim extension As String
        Dim source As String
        Dim count As Integer
        count = 0
        source = InputBox("Enter the source")
        extension = InputBox("Enter the extension of the file")

        If extension.Substring(0, 1) <> "." Then
            extension = String.Concat(".", extension)
        End If
        For Each file1 As String In Directory.GetFiles(source)

            If file1.Contains(extension) Then
                count = count + 1
                If count = 1 Then
                    uploadfiledata_without_trim(file1, "cbs", "Y")
                Else
                    uploadfiledata_without_trim(file1, "cbs", "N")
                End If
            End If

        Next
        For Each Dir As String In Directory.GetDirectories(source)
            uploadsubdirectorydata(Dir, count, extension)
        Next
        MsgBox("Process completed", MsgBoxStyle.Information)
    End Sub
    Sub option607()      '2059 Upload
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        Dim dirs As String() = Directory.GetFiles("c:\du", "*.2059")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        If username.Substring(0, 4).ToUpper <> "MIG0" Then
            MsgBox("Oracle username should start with MIG0")
            Exit Sub
        End If
        If username.Length <> 6 Then
            MsgBox("Invalid Oracle User Name")
            Exit Sub
        End If

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files having extension .2059 exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                uploadfiledata(dir, username, "Y")

                conn.Open()

                processmessage("Inserting data")

                sql = "PKGINSERT2059.INSERT_2059"
                Dim cmd5 As New OracleCommand(sql, conn)
                cmd5.CommandType = CommandType.StoredProcedure
                cmd5.ExecuteNonQuery()

                conn.Close()

            Next

            conn.Open()

            processmessage("Updating FORACID")

            sql = "PKGINSERT2059.UPDATE_FORACID"
            Dim cmd6 As New OracleCommand(sql, conn)
            cmd6.CommandType = CommandType.StoredProcedure
            cmd6.ExecuteNonQuery()

            conn.Close()

            processmessage("")

            MsgBox("Data of " & totalfiles & " files uploaded successfully", MsgBoxStyle.Information, "Process Completed")

            'conn.Dispose()

        End If

    End Sub
    Sub option608()      'Check 2059 files
        Dim lbltext As String
        Dim oracle_mig_string As String = "Data Source=ten;User Id= mig;Password= mig;"
        Dim oracle_mig As New OracleConnection(oracle_mig_string)
        oracle_mig.Open()

        sql = "SELECT MAX(UPLOAD_SRNO) ABCD FROM CID"
        Dim cmd1000 As New OracleCommand(sql, oracle_mig)
        Dim dr1000 As OracleDataReader = cmd1000.ExecuteReader()
        While dr1000.Read
            lbltext = dr1000.Item("ABCD").ToString.Trim
        End While
        dr1000.Close()

        lblinfo10.Text = "Current Customer ID Serial No : " & lbltext
        oracle_mig.Close()

        If username.Substring(0, 4).ToUpper <> "MIG0" Then
            MsgBox("Oracle username should start with MIG0")
            Exit Sub
        End If
        If username.Length <> 6 Then
            MsgBox("Invalid Oracle User Name")
            Exit Sub
        End If

        Dim reportid As String = InputBox("Enter Process ID", "", "ALL")
        Dim mdate As String = InputBox("Enter migration date", "", "01-11-2014")
        ''30-06-2014
        Dim repid As String
        Dim repdesc As String

        Try

            RptDate = CDate(mdate)

        Catch ex As Exception

            MsgBox("Enter valid date", MsgBoxStyle.Critical, "Invalid date")
            Exit Sub

        End Try

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        If reportid = "ALL" Then

            sql = "SELECT SUB_CODE,DESCRIPTION FROM CP WHERE MAIN_CODE = 145 AND SUB_CODE IS NOT NULL ORDER BY SUB_CODE"

        Else

            sql = "SELECT SUB_CODE,DESCRIPTION FROM CP WHERE MAIN_CODE = 145 AND SUB_CODE = '" & reportid & "'"

        End If

        Dim cmd As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd.ExecuteReader()
        While dr.Read
            repid = dr.Item("SUB_CODE").ToString.Trim
            repdesc = dr.Item("DESCRIPTION").ToString.Trim

            processmessage("Checking data - " & repid & " " & repdesc)

            sql = "PKGINSERT2059.CHECK_DATA"
            Dim cmd7 As New OracleCommand(sql, conn)
            cmd7.CommandType = CommandType.StoredProcedure
            cmd7.Parameters.Add("REPORTID", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = repid
            cmd7.Parameters.Add("MDATE", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
            cmd7.ExecuteNonQuery()

            Dim sw1 As StreamWriter = New StreamWriter(folderpath & "\" & repid & " " & repdesc & ".txt")
            Dim sql1 As String = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd1 As New OracleCommand(sql1, conn)
            Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
            While dr1.Read
                Dim linedata As String
                linedata = dr1(0)
                sw1.WriteLine(linedata)
            End While
            dr1.Close()
            sw1.Close()
        End While
        dr.Close()

        conn.Close()

        processmessage("")

        MsgBox("Checking completed successfully", MsgBoxStyle.Information, "Process Completed")


    End Sub

    Public Sub option828()
        Dim source As String
        Dim tablename As String
        Dim new_tablename As String
        Dim new_table_flag As Integer
        Dim delimited_char As String
        Dim delete_existing As Integer
        Dim delete_flag As String = "N"
        Dim new_tab_flag As String = "N"
        Dim lineno As Integer = 1

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)

        source = InputBox("Enter the source folder with full path", "Enter Value", "C:\du")
        tablename = InputBox("Enter the Tablename")
        new_tablename = tablename
        lineno = InputBox("Enter the data starting number (Value 1 or 2)", "Enter Value", "1")
        delete_existing = MsgBox("Do you want to delete existing data from the table and insert  from file?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm")
        If delete_existing = 6 Then
            oracle_execute_non_query("ten", username, username, "DELETE FROM " & tablename)
            delete_flag = "Y"
        End If

        new_table_flag = MsgBox("Do you want to create new table?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm")

        If new_table_flag = 6 Then
            new_tablename = String.Concat(tablename, System.DateTime.Now)
            new_tablename = new_tablename.Replace("/", "")
            new_tablename = new_tablename.Replace("-", "")
            new_tablename = new_tablename.Replace(":", "")
            new_tablename = new_tablename.Replace(" ", "")
            oracle_execute_non_query("ten", username, username, "CREATE TABLE " & new_tablename & " AS SELECT * FROM " & tablename & " WHERE ROWNUM <1")
            new_tab_flag = "Y"
        End If
        delimited_char = InputBox("Enter the Delimited character", "Enter Value", "|")
        Dim count As Integer = 0
        For Each file1 As String In Directory.GetFiles(source)
            count = count + 1
            If count = 1 Then
                uploadfiledata_without_trim(file1, "cbs", "Y")
            Else
                uploadfiledata_without_trim(file1, "cbs", "N")
            End If

        Next
        Thread.Sleep(5000)
        conn.Open()
        sql = "pkgmistool2.INSERTDATA_SPECIFIED_TABLE"
        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.StoredProcedure

        cmd.Parameters.Add("TABLENAME_VAR", OracleDbType.Varchar2, 600, Nothing, ParameterDirection.Input).Value = Trim(new_tablename)
        cmd.Parameters.Add("DEL_FLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = Trim(delete_flag)
        cmd.Parameters.Add("NEW_FLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = Trim(new_tab_flag)
        cmd.Parameters.Add("DELIMITER", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = Trim(delimited_char)
        cmd.Parameters.Add("LINE", OracleDbType.Int16, 10, Nothing, ParameterDirection.Input).Value = lineno
        processmessage("Executing Package....")
        cmd.ExecuteNonQuery()

        conn.Close()
        MsgBox("Process Completed")

    End Sub
    Public Sub uploadsubdirectorydata(ByVal dir, ByVal count, ByVal extension) 'Inserting Subdirectories data into z_du table
        For Each file1 As String In Directory.GetFiles(dir)

            If file1.Contains(extension) Then
                count = count + 1
                If count = 1 Then
                    uploadfiledata_without_trim(file1, "cbs", "Y")
                Else
                    uploadfiledata_without_trim(file1, "cbs", "N")
                End If
            End If

        Next
        For Each dir1 As String In Directory.GetDirectories(dir)
            uploadsubdirectorydata(dir1, count, extension)
        Next
    End Sub

    Private Sub lblinfo8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblinfo8.Click

    End Sub
    Private Sub populate_dgv(ByVal code)
        Dim row As String()
        dgv1.ColumnCount = 2
        dgv1.Rows.Clear()


        If code = 1 Then
            For i As Integer = 0 To menuitems_count
                If menulist(i, 0).ToString.IndexOf(txtcode.Text.ToString()) >= 0 Then
                    row = New String() {menulist(i, 0), menulist(i, 1)}
                    dgv1.Rows.Add(row)
                End If
            Next
        ElseIf code = 2 Then
            For i As Integer = 0 To menuitems_count
                If menulist(i, 1).ToString.ToUpper().IndexOf(txtcode.Text.ToString.ToUpper()) >= 0 Then
                    row = New String() {menulist(i, 0), menulist(i, 1)}
                    dgv1.Rows.Add(row)
                End If
            Next
        End If
    End Sub

    Private Sub format_dgv(ByVal code)
        dgv1.ClearSelection()
        dgv1.ColumnHeadersVisible = False
        dgv1.RowHeadersVisible = False
        dgv1.ReadOnly = True
        Dim column0 As System.Windows.Forms.DataGridViewColumn = dgv1.Columns(0)
        column0.Visible = True
        column0.Width = 60
        Dim column1 As System.Windows.Forms.DataGridViewColumn = dgv1.Columns(1)
        column1.Visible = True
        column1.Width = 615
        dgv1.AllowUserToAddRows = False

    End Sub

    Private Sub txtcode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtcode.KeyDown

        If e.KeyCode = Keys.Enter Then
            If txtmenu.Text = "" Then
                If dgv1.RowCount = 1 Then
                    txtcode.Text = dgv1.CurrentRow.Cells(0).Value
                    txtmenu.Text = dgv1.CurrentRow.Cells(1).Value
                    gprocessid = dgv1.CurrentRow.Cells(0).Value
                    gprocessname = dgv1.CurrentRow.Cells(1).Value
                    dgv1.Visible = False
                    populate_label()
                    Button1.Focus()
                Else
                    dgv1.Visible = True
                    dgv1.Focus()
                End If

            Else
                dgv1.Visible = False
                Button3.Enabled = False
                populate_label()
            End If
        Else

            dgv1.Visible = True

        End If

    End Sub

    Private Sub txtcode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtcode.KeyUp
        dgv1.DataSource = Nothing
        If txtcode.Text.Length = 0 Then
            txtmenu.Text = ""
            clearall()
        End If
        txtmenu.Text = ""

        If e.KeyCode = Keys.Escape Then
            If txtcode.Text.Length = 0 Then
                Application.Exit()
            Else
                txtcode.Text = ""
                clearall()
                txtcode.Focus()
            End If

        End If


        If txtcode.Text.Length >= 2 Then
            dgv1.Visible = True
            If IsNumeric(txtcode.Text) Then
                populate_dgv("1")
                format_dgv("1")
            Else
                populate_dgv("2")
                format_dgv("1")
            End If

        Else

            txtmenu.Text = ""
            clearall()

        End If

        If IsNumeric(txtcode.Text) Then
            txtmenu.Text = ""

            For i As Integer = 0 To menuitems_count
                If txtcode.Text = menulist(i, 0) Then
                    txtmenu.Text = menulist(i, 1)
                    Exit For
                End If
            Next
        End If
    End Sub

    Sub db_open()
        cnn = New OleDb.OleDbConnection
        cnn.ConnectionString = GetConnectionString()
        Try
            If Not cnn.State = ConnectionState.Open Then
                cnn.Open()
            End If
        Catch ex As Exception
            MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
        End Try
    End Sub
    Sub db_close()
        cnn = New OleDb.OleDbConnection
        cnn.ConnectionString = GetConnectionString()
        cnn.Close()
    End Sub

    Private Sub txtmenu_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtmenu.TextChanged

    End Sub

    Private Sub txtcode_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtcode.MouseDown

    End Sub

    Private Sub dgv1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv1.CellClick
        txtcode.Text = ""
        txtmenu.Text = ""
        Try
            txtcode.Text = dgv1.CurrentRow.Cells(0).Value
            txtmenu.Text = dgv1.CurrentRow.Cells(1).Value
            gprocessid = dgv1.CurrentRow.Cells(0).Value
            gprocessname = dgv1.CurrentRow.Cells(1).Value
            dgv1.Visible = False
            populate_label()
        Catch ex As Exception
            txtcode.Text = ""
            txtmenu.Text = ""
            dgv1.Visible = True
        End Try
        Button1.Focus()

    End Sub


    Private Sub populate_label()
        If gprocessname = "Aadhaar Upload - Delete Duplicate Records" Then
            rptoption = 1
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place Aadhaar upload file in c:/du in the name original.txt"
            lblinfo3.Text = "Place Aadhaar error file in c:/du in the name error.txt"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "New Aadhaar upload file will be created in c:/du in the name Aadhaar_New_File.txt"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Daily emails" Then
            rptoption = 2
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40101.email as email.txt and place in c:/du folder"
            lblinfo3.Text = "Ensure that outlook express is running"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Ensure that EMails are generated in the respective outboxes"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "KYC Upload Statistics" Then
            rptoption = 61
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "MASS NEFT AGRICULTURE DEPT" Then
            rptoption = 62
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Process 1 Mass Neft File Creation"
            lblinfo3.Text = "Process 2 Return File Creation"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "kiosk file" Then
            rptoption = 63
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the file in C:\du folder"
            'lblinfo3.Text = "Process 2 Return File Creation"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Email geneared in fag@kgbmis.in outbox"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Weekly Transaction Mail" Then
            rptoption = 64
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the file in C:\du folder"
            'lblinfo3.Text = "Process 2 Return File Creation"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Email geneared ."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Data Upload for DashBoard" Then
            rptoption = 65
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the files(STAFF_NAME.TXT,STAFF_BM.TXT,BUS.TXT) in C:\du folder"
            'lblinfo3.Text = "Process 2 Return File Creation"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Email geneared ."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Upload Files" Then
            rptoption = 3
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the files to be uploaded in c:/du folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Tabdata" Then
            rptoption = 4
            'lblinfo1.Text = "Database: Ten; User Name: CBS"
            lblinfo2.Text = "Place the parameter file in  named as tabdata.txt"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "General" Then
            rptoption = 5
            'lblinfo1.Text = "Database: Ten; User Name: CBS"
            lblinfo2.Text = "Place the parameter file in  named as general.txt"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Report" Then
            rptoption = 6
            'lblinfo1.Text = "Database: Ten; User Name: CBS"
            lblinfo2.Text = "Place the parameter file in  named as report.txt"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "KGB Business Progress Report" Then
            rptoption = 7
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40101.email as 'email.txt' and place in c:/du folder"
            lblinfo3.Text = "Rename the NMGB Trial Balance(TB_XXXXXXXX.xls) as 'nmgb.txt' (Replace tab with |) and place in C:/DU folder"
            lblinfo4.Text = "Rename the NMGB NPA(NPA_XXXXXXXX.xls) File as 'npa.txt' (Replace tab with |) and place in C:/DU folder"
            lblinfo5.Text = "Ensure that the date in all files is similar to that of previous working day"
            lblinfo6.Text = "Run the programme"
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "KGB Day Book" Then
            rptoption = 8
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the NMGB Trial Balance(TB_XXXXXXXX.xls) as 'nmgb.txt' (Replace tab with |) and place in C:/DU folder"
            lblinfo3.Text = "Rename the MISDO File (40124.misdo) as 'smgbdb.txt' and place in C:/DU folder"
            lblinfo4.Text = "Ensure that the date in all files is similar to that of previous working day"
            lblinfo5.Text = "Run the programme"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Business Review" Then
            rptoption = 9
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40102.email as email2.txt and place in c:/du folder"
            lblinfo3.Text = "Enter the previous working days date in 'Previous Working Day' field"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "KGB First - Outstanding" Then
            Button3.Enabled = True
            rptoption = 10
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place SMGB First Outstanding Bank as a whole (MASRPT 802) into C:\DU\ and rename it as 'SMGBFIRST_OS.txt'"
            lblinfo3.Text = "Create a file in c:/DU named MPR_NO_OS.txt reading NMGB File MPR >> Bal_Count after converting to pipe delimited format"
            lblinfo4.Text = "Create a file in c:/DU named MPR_BALANCE_OS.txt reading NMGB File MPR >> Bal_Amt after converting to pipe delimited format"
            lblinfo5.Text = "Run the programme"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "KGB First - Disbursement" Then
            Button3.Enabled = True
            rptoption = 11
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place SMGB First Disbursement Bank as a whole (MASRPT 803) into C:\DU\ and rename it as 'SMGBFIRST_DISB.txt'"
            lblinfo3.Text = "Take branch wise total of eNMGB File MPR >> Disb_Count and Disb_Amt for the month"
            lblinfo4.Text = "Add the figure to total disbursement figure till previous month in Excel File 'MasterFile.xlsx'"
            lblinfo5.Text = "Create a file in c:/DU named MPR_DISB.txt reading the sheet 'Disb_Upto' in Excel file 'MasterFile.xlsx' after converting to pipe delimited format"
            lblinfo6.Text = "Run the programme"
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "KGB First - NPA" Then
            Button3.Enabled = True
            rptoption = 12
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place SMGB First NPA Bank as a whole (MASRPT 805) into C:\DU\ and rename it as 'SMGBFIRST_NPA.txt'"
            lblinfo3.Text = "Create a file in c:/DU named MPR_NO_NPA.txt reading NMGB File MPR >> NPA_Count after converting to pipe delimited format"
            lblinfo4.Text = "Create a file in c:/DU named MPR_BALANCE_NPA.txt reading NMGB File MPR >> NPA_Amt after converting to pipe delimited format"
            lblinfo5.Text = "Run the programme"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "MISDO Upload" Then
            rptoption = 13
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the required MISDO files in C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "After execution of the programme, the data in files '40103-01-10-2013.misdo','40104-01-10-2013.misdo'"
            lblinfo5.Text = "and '40105-01-10-2013.misdo' will be inserted to tables C_BD_GL, C_BD_ADV and C_BD_DEP respectively"
            lblinfo6.Text = "The data in other files, will be inserted to the tables having tablename corresponding the first "
            lblinfo7.Text = "five digit SolID in the file name, prefixed with 'C_'"
            lblinfo8.Text = "For eg: The data in file '40101-01-10-2013.misdo' will be inserted to table C_40101 and so on"
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "ATM Data Mismatch between Finacle & Switch reports" Then
            rptoption = 14
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the Finacle ATM Report as kgbatm_(Date/Period).txt and place in C:/DU folder"
            lblinfo3.Text = "Rename the Switch ATM Report as switchatm.txt and place in C:/DU folder"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "CIBIL Upload File Creation (Live)" Then
            rptoption = 15
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the CIBIL Individual and Non Individual files (Live) generated from Finacle in C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "System will place the output files in D:/CIBIL folder"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "CIBIL Upload File Creation (Close)" Then
            rptoption = 66
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the CIBIL Individual and Non Individual files (Close) generated from Finacle in C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "System will place the output files in D:/CIBIL folder"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Mobile banking SMS creation" Then
            rptoption = 67
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place Mobile banking file from PMO in  C:/DU folder named as SMS.txt"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "System will place the output files in C:/DU folder with name MB_SMS.txt"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Transaction Data upload for DashBoard" Then
            rptoption = 68
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place .6032 files in C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "System will place the output files in C:/DU folder"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "EMail Daily Reports" Then
            rptoption = 16
            lblinfo1.Text = "Generate the files/reports and keep it in c:/temp folder"
            lblinfo2.Text = "Run the programme"
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "NPCI Linked Aadhaar - Upload file creation" Then
            rptoption = 17
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the report downloaded from NPCI in c:/DU naming as npci_aadhaar.txt"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "Output file will be placed in c:/du as npci_aadhaar_upload.txt"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Day end eMails" Then
            rptoption = 18
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the file 40994_XX-XX-XXXX_AC1.TXT as 40994.TXT"
            lblinfo3.Text = "Rename the file 40995_XX-XX-XXXX_AC1.TXT as 40995.TXT"
            lblinfo4.Text = "Rename the file 40998AC1.TXT as 40998.TXT"
            lblinfo5.Text = "Rename the upload error file KYC_XXXXXX.TXT as KYC.TXT"
            lblinfo6.Text = "Rename the file 40991_XX-XX-XXXX_AC1.TXT as 40991.TXT"
            lblinfo7.Text = "Place all files in c:/du folder"
            lblinfo8.Text = "Run the programme"
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Business Review - Files to RO" Then
            rptoption = 19
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the file Business Review Data.txt in '" & Disk & ":\Business Review Report'"
            lblinfo3.Text = "Place the file Business Business Review.docx in '" & Disk & ":\Business Review Report'"
            lblinfo4.Text = "Place the file Business Review.xlsx in '" & Disk & ":\Business Review Report'"
            lblinfo5.Text = "Run the programme"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "KGB Aadhar Enrolled Status" Then
            rptoption = 20
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder"
            lblinfo3.Text = "Create a file in c:/DU named NMGB_AADHAR.txt reading NMGB File AADHARMAPPED.xls after converting to pipe delimited format"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "KGB Daily Reports" Then
            rptoption = 21
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40101_XX-XX-XXXX.email as 'email.txt' and place in c:/du folder"
            lblinfo3.Text = "Rename the NMGB Trial Balance(TB_XXXXXXXX.xls) as 'nmgb.txt' (Replace tab with |) and place in C:/DU folder"
            lblinfo4.Text = "Rename the NMGB NPA(NPA_XXXXXXXX.xls) File as 'npa.txt' (Replace tab with |) and place in C:/DU folder"
            lblinfo5.Text = "Rename the MISDO File (40124_XX-XX-XXXX.misdo) as 'smgbdb.txt' and place in C:/DU folder"
            lblinfo6.Text = "Create a file in c:/DU named NMGB_AADHAR.txt reading NMGB File AADHARMAPPED.xls after converting to pipe delimited format"
            lblinfo7.Text = "Copy the email KGB-BPR send on last friday and replace tab with |, rename it as 'friday.txt' and place in C:/DU folder"
            lblinfo8.Text = "Ensure that the date in all files is similar to that of previous working day"
            lblinfo9.Text = "Run the programme"
            lblinfo10.Text = ""
        ElseIf gprocessname = "9072 Insert" Then
            rptoption = 9072
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place 9072 Files into C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "9074 Insert" Then
            rptoption = 23
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place 9074 Files into C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "9071 Insert" Then
            rptoption = 24
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place 9071 Files into C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Create RO and Branch Folders and convert CIB Files" Then
            rptoption = 25
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Creates folders in c:/du"
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Create Bank as a whole/All RO's/All Branches report in a single file" Then
            rptoption = 26
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensure that the procedure to call exists in PKGMISTOOL3"
            lblinfo3.Text = "Report will be created in C:/DU"
            lblinfo4.Text = "Copy the content to word file"
            lblinfo5.Text = "Replace the word '$$PAGEBREAK$$' with page break"
            lblinfo6.Text = "Reports of 119 Character length can be printed in Letter>>Landscape oreintation (Courier New >> 9)"
            lblinfo7.Text = "Reports of 159 Character length can be printed in Legal>>Landscape oreintation (Courier New >> 9)"
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Get File Names" Then
            rptoption = 27
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Enter the folder path in input box"
            lblinfo3.Text = "Report will be created in C:/DU as FileNames.txt"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Word Document Generation" Then
            rptoption = 28
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Word doc"
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Mobile Banking Transaction Status" Then
            rptoption = 29
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the report of mobile banking transations downloaded from site in C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "Email will be generated in the outbox."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Create Folder" Then
            rptoption = 30
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Specify the folder creation path, By default Sytem will create in c\du folder"
            lblinfo3.Text = "Enter S for SMGB branches N for NMGB branches K for KGB R for RO D for District Folder"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Copy File" Then
            rptoption = 31
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input box1: Enter the file to be copied with full path."
            lblinfo3.Text = "Input box2: Enter the folder in which file to be copied."
            lblinfo4.Text = "Input box3: Copy to subfolders. - Y/N"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Execute Script" Then
            rptoption = 32
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Enter Script file name (with path)"
            lblinfo3.Text = "Enter Access database file name (wihout path)  "
            lblinfo4.Text = "Enter Access database file path (without file name)"
            lblinfo5.Text = "Update in subfolders   - Y/N"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Basedata Generation Timing" Then
            rptoption = 33
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40103_XX-XX-XXXX.email as 'email3.txt' and place in C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "Email will be generated in the outbox."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Staff Upload" Then
            rptoption = 34
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the File in c:/DU "
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "Output file will be placed in c:/du"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "RO Follow Up Status Email" Then
            rptoption = 35
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename one month back email file 40101.email as email_old.txt and place in c:/du folder"
            lblinfo3.Text = "Rename previousday email file 40101.email as email_new.txt and place in c:/du folder"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Email will be generated in the outbox."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "ATM Transaction Status" Then
            rptoption = 36
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the ATM transaction reports in ‘.txt’ format in C:/DU folder."
            lblinfo3.Text = "Run the programme."
            lblinfo4.Text = "Email will be generated in the outbox."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting data into Location table" Then
            rptoption = 801
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Location data transfer"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Data will be inserted into Location table from banc724"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting data into CIDMASTER table" Then
            rptoption = 802
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "CIDMASTER table data transfer"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Data in CEDGE_EXTRACT_CUSTOMERID table will be inserted into CIDMASTER table"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting data to Pickup table" Then
            rptoption = 803
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Pickup data transfer"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Data inserted into Pickup table"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Inserting data to Religioncode table" Then
            rptoption = 804
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Religion Data transfer"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Data inserted into Religioncode table"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Update religioncode from banc724" Then
            rptoption = 805
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating religion code from banc724 data"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Data from banc724 backups updated"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting data to BranchMaster" Then

            rptoption = 806
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating branch master table"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Branch details will be inserted"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting Deposit shadow file" Then

            rptoption = 807
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating AC master table from deposit shadow files"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Input 3 : Enter deposit shadow file path"
            lblinfo6.Text = "Deposit Account numbers inserted into ACMASTER"
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting Loan shadow file" Then

            rptoption = 808
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating AC master table from Loan shadow files"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Input 3 : Enter Loan shadow file path"
            lblinfo6.Text = "Loan Account numbers inserted into ACMASTER"
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Updating NRE code" Then

            rptoption = 809
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating NRE code"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "NRE accounts inserted into NRECODE table"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting Staff Code" Then

            rptoption = 810
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating Staff code"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Staff accounts inserted into STAFFCODE table"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Category code" Then

            rptoption = 811
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating Category Code"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Updated  CATEGORYCODE table from banc724 backup"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting data to Citycode1" Then

            rptoption = 812
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating City Code1"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting data to Citycode2" Then

            rptoption = 813
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating City Code2"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Inserting data to Minor table" Then

            rptoption = 814
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Updating Minor table"
            lblinfo3.Text = "Input 1 : Enter Access database path"
            lblinfo4.Text = "Input 2 : Enter Access database name"
            lblinfo5.Text = "Updated  minor data from banc724 backup"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "uncompress" Then

            rptoption = 815
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place zip file C:\du"
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Inserting Param file and database" Then

            rptoption = 816
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three. Enter the path in which branch folders created."
            lblinfo3.Text = "Enter the access database name with full path."
            lblinfo4.Text = "Param file will be created in each sols Client folder"
            lblinfo5.Text = "IPaddress file and Database file will be created in each sols Server folder."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Copying files for Creating Setup" Then

            rptoption = 817
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three. Enter the (source) path in which Setup folder created(Eg:-C:\du_Setup)"
            lblinfo3.Text = "Enter the  destination folder (Path in which branch folders created)(Eg:- C:\du) "
            lblinfo4.Text = "Enter Y to copy to Subfolders only (Not copied in parent folder). "
            lblinfo5.Text = "Enter Y to copy contents only from the source folder."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "NRE from file" Then

            rptoption = 818
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three. Enter branch folders created path"
            lblinfo3.Text = "Enter access database name"
            lblinfo4.Text = "Update NREcode table from the NRI table"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Deceased from file" Then

            rptoption = 819
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three. Enter branch folders created path"
            lblinfo3.Text = "Enter access database name"
            lblinfo4.Text = "Inserted Deceasecode table from the Deceased table"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Staff no From file" Then

            rptoption = 820
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three.Enter branch folders created path."
            lblinfo3.Text = "Enter the Database name"
            lblinfo4.Text = "Data in stafflist table will be inserted"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Category from file" Then

            rptoption = 821
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three.Enter branch folders created path."
            lblinfo3.Text = "Enter the Database name"
            lblinfo4.Text = "Data in specialcust table will be inserted into Custcategory."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Religion from file" Then

            rptoption = 822
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three.Enter branch folders created path."
            lblinfo3.Text = "Enter the Database name"
            lblinfo4.Text = "Data in Relgion table will be inserted."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Handicapped from file" Then

            rptoption = 823
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three.Enter the path in which the folder created"
            lblinfo3.Text = "Enter the Database name"
            lblinfo4.Text = "Data in handicapped table will be inserted into custcategory"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "LPD details from file" Then

            rptoption = 824
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Ensue user is Three.Enter the path in which the folder created"
            lblinfo3.Text = "Enter the Database name"
            lblinfo4.Text = "Data in LPD1 table will be inserted into LPD table. "
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Compress and email" Then

            rptoption = 825
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Inputbox 1: Enter the folder which is to be Compressed (Source)"
            lblinfo3.Text = "Inputbox 2: Enter the folder where the compressed file kept (Destination)"
            lblinfo4.Text = "The compressed source folder kept in Destination with current date time as name in yyyymmdd_hhmmss format."
            lblinfo5.Text = "Additional Feature: Can mail this Compressed file."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Differential Backup" Then

            rptoption = 826
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input Box1: Enter the folder with full path for creating Backup (Source)."
            lblinfo3.Text = "Input Box2: Enter the folder where the backup kept (Detination)."
            lblinfo4.Text = "Differntial back up of the source folder created and stored in current datetime folder."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Upload - Extension based" Then
            rptoption = 827
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input Box1: Enter the folder where the files kept."
            lblinfo3.Text = "Input Box2: Enter the extension."
            lblinfo4.Text = "The files with the specified extension will be inserted into Z_du table."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Insert into tables" Then
            rptoption = 828
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input Box1: Enter the folder where the files kept. (By default C:\du)"
            lblinfo3.Text = "Input Box2: Enter Table name"
            lblinfo4.Text = "Input Box3: Create new Table.Select Yes/No .If yes new table Created by appednding current datetime"
            lblinfo5.Text = "Input Box4: Enter delimited character."
            lblinfo6.Text = "Data in the Files  will be uploaded into the specified Tables."
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessname = "Differential Backup based on Extension" Then

            rptoption = 829
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input Box1: Enter the folder with full path for creating Backup (Source)."
            lblinfo3.Text = "Input Box2: Enter the folder where the backup kept (Detination)."
            lblinfo4.Text = "Differntial back up of the source folder created and stored in current datetime folder."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Mirror image" Then

            rptoption = 830
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input Box1: Enter the source folder with full path."
            lblinfo3.Text = "Input Box2: Enter the Detination folder with full path."
            lblinfo4.Text = "Source and Destination will be synchronized"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Generating CIDMaster File From dump" Then
            rptoption = 831
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input Box1: Enter Solid"
            lblinfo3.Text = "Input Box2: CIDMaster file will be generated in C:\du folder "
            lblinfo4.Text = "Input Box3: The file name is solid_CIDMASTER.txt format remove solid_ place in  eNMGB Branch folder"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Create text files in a loop" Then
            rptoption = 832
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Write logic in programme source to create text files in the required fashion"
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Create Branch Data" Then
            rptoption = 601
            lblinfo2.Text = "Input Box1: Enter Branch Code."
            lblinfo3.Text = "Input Box2: Enter Username in which the dump imported."
            lblinfo4.Text = "Create a folder " & Disk & ":\dump\script.Place the following files."
            lblinfo5.Text = "droptenFUF_USERS.sql,Script_tomove_data.sql,Table_script.sql"
            lblinfo6.Text = "Package_script.sql,Update_script.sql"
            lblinfo7.Text = "Create a folder " & Disk & ":\dump\BR_solid(In five digit eg:-BR_00001)."
            lblinfo8.Text = "Place the Bsolid.dmp,Tsolid.dmp in " & Disk & ":\dump\BR_solid."
            lblinfo9.Text = ""
            lblinfo10.Text = ""


        ElseIf gprocessname = "eNMGB Migration - Upload Migration Tool Files" Then
            rptoption = 602
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the files solid_816_101 to solid_816_115 in c:\du folder"
            lblinfo3.Text = "All data will be inserted into Corresponding MT tables."
            lblinfo4.Text = "Summary File 301_Data_Entry_Status  will be generated in C:\du folder"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Upload CGL File" Then
            rptoption = 603
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the files username_cgl.txt in c:\du folder"
            lblinfo3.Text = "All data will be inserted into MT_cgl Table."
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Assign CustID and Account No" Then
            rptoption = 604
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the latest CID.txt file in C:/DU"
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - FUF Generation" Then
            rptoption = 605
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input Box1: Enter Report ID"
            lblinfo3.Text = "Input Box2: Enter the date of migartion"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Reports will be generated in two folders: C:\DU\Branch Folder\Report ID Folder\MIS Team & C:\DU\Branch Folder\Report ID Folder\Wipro"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Reports" Then
            rptoption = 606
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Input Box1: Enter Branch Code"
            lblinfo3.Text = "Input Box2: Enter Report ID"
            lblinfo4.Text = "Input Box3: Reports will be generated in: C:\DU\Branch Folder\Report ID Folder"
            lblinfo5.Text = "Input Box4: The file name is BranchCode_ReportID_Reportfilename.txt format."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Migration Tool Data Entry Status Email" Then
            rptoption = 39
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the Migration Tool data entry status files in 'c:\du' folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Export Oracle Data" Then
            rptoption = 40
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Backup, Drop and Import Oracle Tables" Then

            rptoption = 41
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Drop oracle user" Then

            rptoption = 42
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Figures At A Glance" Then

            rptoption = 43
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40101.email as email.txt and place in c:/du folder"
            lblinfo3.Text = "Ensure that outlook express is running"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Ensure that all eMails are generated in the respective outboxes"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "PMJDY Campaign" Then

            rptoption = 45
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Run the programme"
            lblinfo3.Text = "Ensure that all eMails are generated in the respective outboxes"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Business Figures As On 30-09-2014" Then

            rptoption = 47
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Run the programme"
            lblinfo3.Text = "Ensure that all eMails are generated in the respective outboxes"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Branch Intimation Letter" Then

            rptoption = 48
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Run the programme"
            lblinfo3.Text = "Ensure that all eMails are generated in the respective outboxes"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

            'ElseIf gprocessname = "SARFAESI Notice Intimation Status" Then
        ElseIf gprocessname = "PMJJBY/PMSBY/APY Enrollment Status" Then

            rptoption = 53
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the PMJJBY/PMSBY data file in C:/DU folder as 'Data.txt'. Ensure that C:/DU folder has only the mentioned file."
            lblinfo3.Text = "Place the APY data file in C:/DU folder as 'Data1.txt'. Ensure that C:/DU folder has only the mentioned file."
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Ensure that all eMails are generated in the respective outboxes."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "NPA Threat For Next 7 Days - Email Generation" Then

            rptoption = 55
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Process to be run after creation of excel files."
            lblinfo3.Text = "Run the programme."
            lblinfo4.Text = "Ensure that all eMails are generated in the respective outboxes."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "NPA Threat For Next 7 Days - Excel Creation" Then

            rptoption = 54
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Create a folder 'PNPA' in D: folder."
            lblinfo3.Text = "Place all .9114 files in C:/DU folder. Ensure that C:/DU folder has only the mentioned file."
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Ensure that 44 Excel files are generated in the D:/PNPA folder."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "NPA Threat For Next 7 Days - Excel Creation Using Macro" Then

            rptoption = 56
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Create a folder 'PNPA' in D: folder."
            lblinfo3.Text = "Place all .9114 files in C:/DU folder. Ensure that C:/DU folder has only the mentioned file."
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Ensure that 44 Excel files are generated in the D:/PNPA folder."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Predefined Day End Check Validation" Then

            rptoption = 57
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place all .9114 files in C:/DU folder. Ensure that C:/DU folder has only the mentioned file."
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "BOD Mails" Then

            rptoption = 58
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place .9088 files in C:/DU Folder"
            lblinfo3.Text = "Place last friday figure file as FRIDAY.txt in C:/Du Folder"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Ensure that all eMails are generated in the respective outboxes."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "NPA Reports" Then

            rptoption = 59
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place Mydump.dmp in D Drive"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "Accept Defualt path and Default file name"
            lblinfo5.Text = "Enter ALL to execute the all program/GENERATE EMAIL or GENERATED REPORTS to execute the part only."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "PROGRESS REPORT AS PER CIRCULAR: 74/2015" Then

            rptoption = 60
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place .9094 files in C:/DU folder. Ensure that C:/DU folder has only the mentioned file."
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "Ensure that all eMails are generated in the respective outboxes."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Upload 2059 Files" Then
            rptoption = 607
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the required 2059 files in C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Check 2059 Files" Then
            rptoption = 608
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Should execute only after uploading 2509 files through Option No - 607"
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Split CEDGE Dump" Then
            rptoption = 609
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Create History Transaction Data Dump" Then
            rptoption = 610
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Create NPA Upload Files" Then
            rptoption = 611
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Batch update of packages" Then
            rptoption = 612
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Create backup of live users" Then
            rptoption = 613
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Import Users" Then
            rptoption = 614
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Data from users" Then
            rptoption = 615
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "eNMGB Migration - Zenith Backup Import" Then
            rptoption = 616
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf gprocessname = "Bulk SMS File Creation" Then
            rptoption = 46
            lblinfo1.Text = "Designation Input Options : ALL,CH,GM,RM,SM,MG(For chairman and GM enter CH,GM)"
            lblinfo2.Text = "Office Type : ALL,HO,RO,BR,HD  (IF Office type is All Message generated for Braches also)"
            lblinfo3.Text = "Department : ALL,CW,CS,HW,IT,RL,PD"
            lblinfo4.Text = "smsupd.txt file will be generated in c:\temp folder"
            lblinfo5.Text = "Upload and confirm in live Server through BATCHUPD >> SMSBATCH"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf gprocessid = "617" Then
            rptoption = 617
            lblinfo1.Text = "Write procedure in PKGLAPTOP20SESSION"
            lblinfo2.Text = ""
            lblinfo3.Text = ""
            lblinfo4.Text = ""
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        End If

        '' FRANKLIN - UPDATE LABEL INFO
        '' NEXT CREATE OPTIONXX FUNCTION

        For i As Integer = 0 To menuitems_count
            If menulist(i, 0) = txtcode.Text Then
                If menulist(i, 2) = "D" Then
                    MsgBox("This option is not available now. Currently in deleted status")
                End If
            End If

        Next

    End Sub

    Private Sub dgv1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv1.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            txtcode.Text = ""
            txtmenu.Text = ""
            Try
                txtcode.Text = dgv1.CurrentRow.Cells(0).Value
                txtmenu.Text = dgv1.CurrentRow.Cells(1).Value
                gprocessid = dgv1.CurrentRow.Cells(0).Value
                gprocessname = dgv1.CurrentRow.Cells(1).Value
                dgv1.Visible = False
                populate_label()
                MsgBox("You have selected Option ID - " & txtcode.Text & " - " & txtmenu.Text)
            Catch ex As Exception
                txtcode.Text = ""
                txtmenu.Text = ""
                dgv1.Visible = True
                MsgBox("Please select a valid Report ID")
            End Try
            'Thread.Sleep(2000)


        End If
    End Sub

    Private Sub copyfile(ByVal source, ByVal dest, ByVal temp)
        For Each dir1 As String In Directory.GetDirectories(temp)
            Dim temp1 As String
            temp1 = ""
            temp1 = temp
            copyfile(source, dest, dir1)
            temp1 = Path.Combine(temp, Path.GetFileName(source))
            File.Copy(source, temp1, True)
        Next
        temp = Path.Combine(temp, Path.GetFileName(source))
        File.Copy(source, temp, True)
    End Sub

    Private Sub executescriptInsubfolder(ByVal scriptfile, ByVal exepath, ByVal exedb)

        If File.Exists(exepath & "\" & exedb) Then
            Dim filevar As String = scriptfile
            Dim Line As String = ""


            'If File.Exists(exepath & "\" & exedb) = False Then
            '    MsgBox("Cannot find the Destination file,Please check", MsgBoxStyle.Critical)
            '    Exit Sub
            'End If

            Dim cnn As New OleDb.OleDbConnection
            cnn = New OleDb.OleDbConnection

            Dim strConnection As String
            strConnection = ""
            strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & exepath & "\" & exedb
            cnn.ConnectionString = strConnection

            Try
                If Not cnn.State = ConnectionState.Open Then
                    cnn.Open()
                End If
            Catch ex As Exception
                MsgBox("Cannot find/open database", MsgBoxStyle.Critical, "Cannot Proceed")
                MsgBox(exepath & "-" & exedb)
            End Try

            If System.IO.File.Exists(filevar) = True Then
                processmessage("Executing script in file " & exepath & "\" & exedb)
                Dim objReader As New System.IO.StreamReader(filevar)
                Do While objReader.Peek() <> -1
                    Line = ""
                    Line = Line & objReader.ReadLine() & vbNewLine
                    '    Line = Line.Remove(0, 2)
                    'Line = readNthLine(filevar, 0)
                    Try
                        Dim cmd As New OleDb.OleDbCommand
                        cmd.CommandText = Line
                        cmd.Connection = cnn
                        cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox("An error was raised!" & vbNewLine & "Message: " & Err.Description, MsgBoxStyle.Critical, "Error")
                        'cnn.Close()
                        'Exit Sub
                    End Try
                Loop
                objReader.Close()
                cnn.Close()
            End If
        End If
        For Each dir1 In Directory.GetDirectories(exepath)
            executescriptInsubfolder(scriptfile, dir1, exedb)
        Next

    End Sub
    Private Sub clearall()
        'txtcode.Text = ""
        txtmenu.Text = ""
        lblstatus.Text = ""
        lblstatus2.Text = ""
        lblinfo1.Text = ""
        lblinfo2.Text = ""
        lblinfo3.Text = ""
        lblinfo4.Text = ""
        lblinfo5.Text = "Enter Option"
        lblinfo6.Text = ""
        lblinfo7.Text = ""
        lblinfo8.Text = ""
        lblinfo9.Text = ""
        lblinfo10.Text = ""
        dgv1.Visible = False
        txtcode.Focus()


    End Sub
    Public Function GetConnectionString() As String
        Dim strConnection As String
        Dim directory As String = My.Application.Info.DirectoryPath & "\MISTOOL.mdb"
        strConnection = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & directory
        GetConnectionString = strConnection
    End Function

    Public Function CalculateEMI() As Double


        Dim scheduledate As Date
        Dim scheduleamount As Double
        Dim scheduleperiod As Double
        Dim schedulerepayoption As String
        Dim schedulerepayfrequency As String
        Dim schedulerph As Double
        Dim capitaliserphint As String
        Dim txtnewrate As String
        Dim txtoldinterest As String
        Dim oldinterest As Double
        Dim newrate As Double

        Dim lastinstallmentdate As Date
        Dim installmentdivisor As Double
        Dim numberofinstallments As Double
        Dim numberofmonths As Integer
        Dim oddinstallmentnumber As String
        Dim scheduleinterest As Double
        Dim rphcompletiondate As Date
        Dim txtnoofinstallments As Integer
        Dim intonrph As Double
        Dim txtinttocapitalize As Double
        Dim txtschamt As Double
        Dim oldrateperiod As Integer
        Dim newrateperiod As Integer
        Dim numberofmonths2 As Integer
        Dim Serialno As Integer = 0

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        oracle_execute_non_query("ten", username, username, "DELETE FROM FUF_LRS_1")
        sql = "SELECT F_ACNO,SDATE,SAMOUNT,STERM,NEW_SREPAYTYPE,SREPAYFREQ,SRPH,SRPHINTCAPITALIZE,OLDINTEREST,CURRINTEREST FROM FUF_LRS WHERE RECORD_FLAG IN ('CEDGE','MT')"
        Dim cmd101 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd101.ExecuteReader()
        While dr.Read()
            Serialno = Serialno + 1
            acno = dr.Item("F_ACNO")
            scheduledate = dr.Item("SDATE")
            scheduleamount = dr.Item("SAMOUNT")
            scheduleperiod = dr.Item("STERM")
            schedulerepayoption = dr.Item("NEW_SREPAYTYPE")
            schedulerepayfrequency = dr.Item("SREPAYFREQ")
            schedulerph = dr.Item("SRPH")
            capitaliserphint = dr.Item("SRPHINTCAPITALIZE")
            txtnewrate = dr.Item("CURRINTEREST")
            txtoldinterest = dr.Item("OLDINTEREST")
            oldinterest = txtoldinterest
            newrate = txtnewrate
            'overdueason = "22-05-2014"
            dtintchangedate = "01-09-2013"
            lastinstallmentdate = scheduledate.AddMonths(scheduleperiod)

            If schedulerepayfrequency = "01" Then
                installmentdivisor = 1
            ElseIf schedulerepayfrequency = "03" Then
                installmentdivisor = 3
            ElseIf schedulerepayfrequency = "06" Then
                installmentdivisor = 6
            ElseIf schedulerepayfrequency = "12" Then
                installmentdivisor = 12
            End If

            numberofinstallments = (scheduleperiod - schedulerph) / installmentdivisor
            numberofmonths = scheduleperiod - schedulerph

            If Math.Floor(numberofinstallments) <> numberofinstallments Then
                numberofinstallments = Math.Floor(numberofinstallments) + 1
                oddinstallmentnumber = "Y"
            Else
                oddinstallmentnumber = "N"
            End If

            If txtoldinterest = "0" Then
                scheduleinterest = txtnewrate
            Else
                scheduleinterest = txtoldinterest
            End If

            'accountbalance = txtacbalance.Text
            rphcompletiondate = scheduledate.AddMonths(schedulerph)

            txtnoofinstallments = numberofinstallments  ''NO OF INSTALLMENTS

            If schedulerepayoption <> "1" Then

                If capitaliserphint = "Y" Then
                    intonrph = Math.Round(scheduleamount * schedulerph * scheduleinterest / 1200, 2)
                    scheduleamount = scheduleamount + intonrph
                    txtinttocapitalize = intonrph
                    txtschamt = scheduleamount
                Else
                    txtinttocapitalize = "0"
                    txtschamt = scheduleamount
                End If

                'If scheduleamount > Val(txtloanamount.Text) Then
                '    MsgBox("For non EI loans, schedule amount should not exceed sanction limit. If rescheduled/rephased, enter the schedule details in the rescheduled/rephased part in a way that the total schedule amount do not exceed the loan amount", MsgBoxStyle.Information, "Attention")
                '    Exit Function
                'End If

                generate_schedule(scheduleamount, numberofmonths, scheduleinterest, schedulerepayfrequency, schedulerepayoption, rphcompletiondate, "N")

            Else

                If (oldinterest = "0" Or oldinterest = newrate Or scheduledate >= dtintchangedate) Then

                    If capitaliserphint = "Y" Then
                        intonrph = Math.Round(scheduleamount * schedulerph * scheduleinterest / 1200, 2)
                        scheduleamount = scheduleamount + intonrph
                        txtinttocapitalize = intonrph
                        txtschamt = scheduleamount
                    Else
                        txtinttocapitalize = "0"
                        txtschamt = scheduleamount
                    End If

                    generate_schedule(scheduleamount, numberofmonths, scheduleinterest, schedulerepayfrequency, schedulerepayoption, rphcompletiondate, "N")

                ElseIf lastinstallmentdate < dtintchangedate Then

                    If capitaliserphint = "Y" Then
                        intonrph = Math.Round(scheduleamount * schedulerph * scheduleinterest / 1200, 2)
                        scheduleamount = scheduleamount + intonrph
                        txtinttocapitalize = intonrph
                        txtschamt = scheduleamount
                    Else
                        txtinttocapitalize = "0"
                        txtschamt = scheduleamount
                    End If

                    generate_schedule(scheduleamount, numberofmonths, oldinterest, schedulerepayfrequency, schedulerepayoption, rphcompletiondate, "N")

                ElseIf scheduledate < dtintchangedate And rphcompletiondate > dtintchangedate And capitaliserphint = "Y" Then

                    oldrateperiod = DateDiff(DateInterval.Day, scheduledate, dtintchangedate)
                    newrateperiod = DateDiff(DateInterval.Day, dtintchangedate, rphcompletiondate)
                    intonrph = Math.Round(scheduleamount * oldrateperiod * oldinterest / 36500, 2)
                    intonrph = intonrph + (Math.Round(scheduleamount * newrateperiod * newrate / 36500, 2))
                    scheduleamount = scheduleamount + intonrph
                    txtinttocapitalize = intonrph
                    txtschamt = scheduleamount
                    generate_schedule(scheduleamount, numberofmonths, newrate, schedulerepayfrequency, schedulerepayoption, rphcompletiondate, "N")

                ElseIf scheduledate < dtintchangedate And rphcompletiondate > dtintchangedate And capitaliserphint <> "Y" Then

                    txtinttocapitalize = 0
                    txtschamt = scheduleamount
                    generate_schedule(scheduleamount, numberofmonths, newrate, schedulerepayfrequency, schedulerepayoption, rphcompletiondate, "N")

                ElseIf dtintchangedate >= rphcompletiondate And capitaliserphint = "Y" Then

                    intonrph = Math.Round(scheduleamount * schedulerph * oldinterest / 1200, 2)
                    scheduleamount = scheduleamount + intonrph
                    txtinttocapitalize = intonrph
                    txtschamt = scheduleamount

                    generate_schedule(scheduleamount, numberofmonths, oldinterest, schedulerepayfrequency, schedulerepayoption, rphcompletiondate, "Y")
                    numberofmonths2 = numberofmonths - int_change_date_no_of_months
                    If numberofmonths2 > 0 And int_change_date_theo_balance > 0 Then
                        generate_schedule(int_change_date_theo_balance, numberofmonths2, newrate, schedulerepayfrequency, schedulerepayoption, int_change_date_inst_date, "N")
                    Else
                        MsgBox("Number of months and interest change date theo balance is zero for second schedule for account no - " & acno)
                    End If

                ElseIf dtintchangedate >= rphcompletiondate And capitaliserphint <> "Y" Then

                    txtinttocapitalize = 0
                    txtschamt = scheduleamount

                    generate_schedule(scheduleamount, numberofmonths, oldinterest, schedulerepayfrequency, schedulerepayoption, rphcompletiondate, "Y")
                    numberofmonths2 = numberofmonths - int_change_date_no_of_months
                    If numberofmonths2 > 0 And int_change_date_theo_balance > 0 Then
                        generate_schedule(int_change_date_theo_balance, numberofmonths2, newrate, schedulerepayfrequency, schedulerepayoption, int_change_date_inst_date, "N")
                    Else
                        MsgBox("Number of months and interest change date theo balance is zero for second schedule for account no - " & acno)
                    End If
                Else

                    If capitaliserphint = "Y" Then
                        intonrph = Math.Round(scheduleamount * schedulerph * scheduleinterest / 1200, 2)
                        scheduleamount = scheduleamount + intonrph
                        txtinttocapitalize = intonrph
                        txtschamt = scheduleamount
                    Else
                        txtinttocapitalize = "0"
                        txtschamt = scheduleamount
                    End If

                    generate_schedule(scheduleamount, numberofmonths, scheduleinterest, schedulerepayfrequency, schedulerepayoption, rphcompletiondate, "N")

                End If

            End If

            oracle_execute_non_query("ten", username, username, "UPDATE FUF_LRS SET INSTALLMENT = " & txtemi & ", NO_OF_INSTALLMENTS = " & txtnoofinstallments & ", INT_CAPITALIZED = " & txtinttocapitalize & " WHERE F_ACNO = '" & acno & "'")
            processmessage("Generating schedule for account no - " & Serialno)
        End While
        dr.Close()

        CalculateEMI = txtemi  '' EMI

    End Function

    Sub generate_schedule(ByVal scheduleamount, ByVal numberofmonths, ByVal scheduleinterest, ByVal installmentfrequency, ByVal repaytype, ByVal startdate, ByVal StopOnIntChange)

        Dim eff_int_rate As Double
        Dim no_of_installments As Double
        Dim eff_installment_no As Double
        Dim inst_amt As Double

        Dim prin_bal As Double
        Dim prin_part As Double
        Dim int_part As Double
        Dim inst_date As Date
        Dim theo_balance As Double


        Dim row As String()

        inst_date = startdate
        int_change_date_theo_balance = 0
        int_change_date_no_of_months = 0

        If installmentfrequency = "01" Then
            eff_int_rate = scheduleinterest / 100 / 12
            no_of_installments = numberofmonths
        ElseIf installmentfrequency = "03" Then
            eff_int_rate = scheduleinterest / 100 / 4
            no_of_installments = numberofmonths / 3
        ElseIf installmentfrequency = "06" Then
            eff_int_rate = scheduleinterest / 100 / 2
            no_of_installments = numberofmonths / 6
        ElseIf installmentfrequency = "12" Then
            eff_int_rate = scheduleinterest / 100
            no_of_installments = numberofmonths / 12
        End If

        eff_installment_no = Math.Round(no_of_installments + 0.499)
        prin_bal = scheduleamount

        If repaytype = "1" Then

            inst_amt = Math.Round(Pmt(eff_int_rate, no_of_installments, -scheduleamount), 2)

            For loopcount = 1 To eff_installment_no

                If loopcount = eff_installment_no Or loopcount > eff_installment_no Then
                    int_part = Math.Round(IPmt(eff_int_rate, eff_installment_no, no_of_installments, -scheduleamount) * (no_of_installments - (eff_installment_no - 1)), 2)
                    prin_part = Math.Round(prin_bal, 2)
                Else
                    int_part = Math.Round(IPmt(eff_int_rate, loopcount, no_of_installments, -scheduleamount), 2)
                    prin_part = Math.Round(inst_amt - int_part, 2)
                End If

                prin_bal = Math.Round(prin_bal - prin_part, 2)

                If installmentfrequency = "01" Then
                    inst_date = inst_date.AddMonths(1)
                ElseIf installmentfrequency = "03" Then
                    inst_date = inst_date.AddMonths(3)
                ElseIf installmentfrequency = "06" Then
                    inst_date = inst_date.AddMonths(6)
                ElseIf installmentfrequency = "12" Then
                    inst_date = inst_date.AddMonths(12)
                End If

                If loopcount = eff_installment_no Or loopcount > eff_installment_no Then
                    If inst_date > startdate.AddMonths(numberofmonths) Then
                        inst_date = startdate.AddMonths(numberofmonths)
                    End If
                End If

                If inst_date <= overdueason Then
                    theo_balance = prin_bal
                    txtnewod = Math.Round((accountbalance - theo_balance), 2)
                End If

                If StopOnIntChange = "N" Then

                    row = New String() {loopcount, inst_date, prin_part, int_part, prin_part + int_part, prin_bal}
                    oracle_execute_non_query("ten", username, username, "INSERT INTO FUF_LRS_1 (F_ACNO,INST_NO,INST_DATE,PRIN_PART,INT_PART,TOTAL_INST,PRIN_BAL) VALUES('" & acno & "','" & loopcount & "',TO_DATE('" & inst_date & "','DD-MM-YYYY'),'" & prin_part & "','" & int_part & "','" & prin_part + int_part & "','" & prin_bal & "')")

                Else

                    If inst_date <= dtintchangedate Then

                        row = New String() {loopcount, inst_date, prin_part, int_part, prin_part + int_part, prin_bal}
                        oracle_execute_non_query("ten", username, username, "INSERT INTO FUF_LRS_1 (F_ACNO,INST_NO,INST_DATE,PRIN_PART,INT_PART,TOTAL_INST,PRIN_BAL) VALUES('" & acno & "','" & loopcount & "',TO_DATE('" & inst_date & "','DD-MM-YYYY'),'" & prin_part & "','" & int_part & "','" & prin_part + int_part & "','" & prin_bal & "')")
                        int_change_date_theo_balance = prin_bal
                        int_change_date_inst_date = inst_date

                        If installmentfrequency = "01" Then
                            int_change_date_no_of_months = loopcount
                        ElseIf installmentfrequency = "03" Then
                            int_change_date_no_of_months = loopcount * 3
                        ElseIf installmentfrequency = "06" Then
                            int_change_date_no_of_months = loopcount * 6
                        ElseIf installmentfrequency = "12" Then
                            int_change_date_no_of_months = loopcount * 12
                        End If

                    ElseIf loopcount = 1 And inst_date > dtintchangedate Then

                        row = New String() {loopcount, inst_date, prin_part, int_part, prin_part + int_part, prin_bal}
                        oracle_execute_non_query("ten", username, username, "INSERT INTO FUF_LRS_1 (F_ACNO,INST_NO,INST_DATE,PRIN_PART,INT_PART,TOTAL_INST,PRIN_BAL) VALUES('" & acno & "','" & loopcount & "',TO_DATE('" & inst_date & "','DD-MM-YYYY'),'" & prin_part & "','" & int_part & "','" & prin_part + int_part & "','" & prin_bal & "')")
                        int_change_date_theo_balance = prin_bal
                        int_change_date_inst_date = inst_date

                        If installmentfrequency = "01" Then
                            int_change_date_no_of_months = loopcount
                        ElseIf installmentfrequency = "03" Then
                            int_change_date_no_of_months = loopcount * 3
                        ElseIf installmentfrequency = "06" Then
                            int_change_date_no_of_months = loopcount * 6
                        ElseIf installmentfrequency = "12" Then
                            int_change_date_no_of_months = loopcount * 12
                        End If

                    End If

                End If

            Next

        Else

            inst_amt = Math.Round(scheduleamount / eff_installment_no, 2)

            For loopcount = 1 To eff_installment_no

                int_part = 0
                prin_part = inst_amt
                prin_bal = Math.Round(prin_bal - prin_part, 2)

                If installmentfrequency = "01" Then
                    inst_date = inst_date.AddMonths(1)
                ElseIf installmentfrequency = "03" Then
                    inst_date = inst_date.AddMonths(3)
                ElseIf installmentfrequency = "06" Then
                    inst_date = inst_date.AddMonths(6)
                ElseIf installmentfrequency = "12" Then
                    inst_date = inst_date.AddMonths(12)
                End If

                If loopcount = eff_installment_no Or loopcount > eff_installment_no Then
                    If inst_date > startdate.AddMonths(numberofmonths) Then
                        inst_date = startdate.AddMonths(numberofmonths)
                    End If
                End If

                If inst_date <= overdueason Then
                    theo_balance = prin_bal
                    txtnewod = Math.Round((accountbalance - theo_balance), 2)
                End If

                row = New String() {loopcount, inst_date, prin_part, int_part, prin_part + int_part, prin_bal}
                oracle_execute_non_query("ten", username, username, "INSERT INTO FUF_LRS_1 (F_ACNO,INST_NO,INST_DATE,PRIN_PART,INT_PART,TOTAL_INST,PRIN_BAL) VALUES('" & acno & "','" & loopcount & "',TO_DATE('" & inst_date & "','DD-MM-YYYY'),'" & prin_part & "','" & int_part & "','" & prin_part + int_part & "','" & prin_bal & "')")

            Next

        End If

        txtemi = inst_amt

    End Sub

    Private Sub WorkingProcess()
        Dim myProcesses() As Process
        Dim myProcess As Process

        Dim inprocessing As Integer = 0
        myProcesses = Process.GetProcesses()

        inprocessing = 0
        For Each myProcess In myProcesses
            If myProcess.ProcessName = "cmd" Then
                inprocessing = 1
                Exit For
            End If
            If myProcess.ProcessName = "imp" Then
                inprocessing = 1
                Exit For
            End If
        Next

        If inprocessing = 1 Then
            flag = 1
        Else
            flag = 0
        End If
    End Sub
    Private Sub Compress_gzip(ByVal fi As FileInfo, ByVal deleteflag As String)
        Dim filename As String = fi.FullName
        'If filename.Contains(".") Then
        '    'MsgBox((filename.IndexOf(".") + 1))
        '    filename = filename.Remove(filename.IndexOf("."), filename.Length - filename.IndexOf("."))
        '    'MsgBox(filename)
        'End If
        Dim pasta As String = filename

        Dim fs As FileStream = New FileStream(fi.ToString(), FileMode.Open)
        Dim input(fs.Length) As Byte
        fs.Read(input, 0, input.Length)
        fs.Close()

        Dim fsOutput As FileStream = New FileStream(pasta + ".gz", FileMode.Create, FileAccess.Write)
        Dim zip As GZipStream = New GZipStream(fsOutput, CompressionMode.Compress)

        zip.Write(input, 0, input.Length)
        zip.Close()
        fsOutput.Close()

        If deleteflag.ToUpper = "Y" Then
            File.Delete(fi.ToString())
        End If
    End Sub

    Private Sub option616()   'eNMGB Migration - Zenith Backup Import

        Dim dumppath As String = InputBox("Enter path in which is backups are placed", "Enter Value", "D:\1")
        If dumppath = "" Then
            MsgBox("No data entered.  Exiting application")
            Exit Sub
        End If

        Dim tempcount As Integer
        Dim processrunning As Boolean
        Dim myProcesses() As Process
        Dim myProcess As Process

        Dim dirs As String() = Directory.GetFiles(dumppath, "*.dmp")
        Dim dir As String
        Dim totalfiles As Integer
        Dim filename As String
        Dim uname As String

        totalfiles = dirs.Length

        If totalfiles = 0 Then
            MsgBox("No .dmp exists in the folder " & dumppath, MsgBoxStyle.Critical, "Error")
        Else
            For Each dir In dirs
                tempcount = tempcount + 1
                filename = GetFileName(dir)
                uname = "three"

                Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
                sw0.WriteLine("@echo off")
                sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
                sw0.Close()

                Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
                sw.WriteLine("connect three/three;")
                sw.WriteLine("drop table dosys010;")
                sw.WriteLine("drop table domst020;")
                sw.WriteLine("drop table dotrn020;")
                sw.WriteLine("exit")
                sw.Close()

                Process.Start(Disk & ":\dump\script\create_user.bat")

                processrunning = True
                While processrunning
                    tempcount = 0
                    myProcesses = Process.GetProcesses()
                    For Each myProcess In myProcesses
                        If UCase(myProcess.ProcessName) = "CMD" Then
                            tempcount = 1
                        End If
                    Next
                    If tempcount = 0 Then
                        processrunning = False
                    End If
                    If processrunning = True Then
                        Thread.Sleep(2000)
                    End If
                End While

                Dim sw1 As StreamWriter = New StreamWriter(Disk & ":\dump\script\import_user_br.bat")
                sw1.WriteLine("@echo off")
                sw1.WriteLine("imp " & uname & "/" & uname & "@ten file=" & dir & " TABLES=(domst020, dosys010, dotrn020) GRANTS=N CONSTRAINTS=N INDEXES=N IGNORE=Y")
                sw1.Close()
                Process.Start(Disk & ":\dump\script\import_user_br.bat")

                processrunning = True
                While processrunning
                    tempcount = 0
                    myProcesses = Process.GetProcesses()
                    For Each myProcess In myProcesses
                        If UCase(myProcess.ProcessName) = "CMD" Then
                            tempcount = 1
                        End If
                    Next
                    If tempcount = 0 Then
                        processrunning = False
                    End If
                    If processrunning = True Then
                        Thread.Sleep(2000)
                    End If
                End While

                Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
                sw2.WriteLine("@echo off")
                sw2.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
                sw2.Close()

                Dim sw3 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
                sw3.WriteLine("connect three/three;")
                sw3.WriteLine("INSERT INTO DOSYS010_TEMP (BRCP_BANK_CODE,BRCP_BRANCH_CODE,BRCP_BANK_NAME,BRCP_BRANCH_NAME,ADD_DATE) SELECT BRCP_BANK_CODE,BRCP_BRANCH_CODE,BRCP_BANK_NAME,BRCP_BRANCH_NAME,SYSTIMESTAMP FROM DOSYS010;")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("INSERT INTO DOMST020_TEMP SELECT ACCT_BANK_CD,ACCT_BRANCH_CD,ACCT_SUBSYS_CD,ACCT_AC_NO,ACCT_GLOBAL_ID,ACCT_NAME,ACCT_AC_OPEN_DATE,ACCT_CLEAR_BAL,ACCT_TOTAL_BAL,ACCT_STATUS FROM DOMST020 WHERE ACCT_SUBSYS_CD IN ('OSLEDL','DLEDL') AND ACCT_STATUS NOT IN (98) AND ACCT_AC_CLOSED_DATE IS NULL;")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("INSERT INTO DOMST020_MASTER SELECT * FROM DOMST020_TEMP;")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("INSERT INTO DOTRN020_TEMP SELECT MTRN_BANK_CD,MTRN_BRANCH_CD,MTRN_DATE,MTRN_SCROLL_NO_X,MTRN_SCROLL_NO_9,MTRN_VOUCHER_SRNO,MTRN_SUBSYS_CD,MTRN_AC_NO,MTRN_ASOF_DATE,MTRN_INS_TYPE,MTRN_AMOUNT,MTRN_PARTICULARS,MTRN_STATUS FROM DOTRN020 WHERE (MTRN_SUBSYS_CD,MTRN_AC_NO) IN (SELECT ACCT_SUBSYS_CD,ACCT_AC_NO FROM DOMST020_TEMP);")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("INSERT INTO DOTRN020_MASTER SELECT * FROM DOTRN020_TEMP;")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("UPDATE DOSYS010_TEMP SET MAX_TRAN_DATE = (SELECT MAX(MTRN_DATE) FROM DOTRN020);")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("UPDATE DOSYS010_TEMP SET AC_COUNT = (SELECT COUNT(1) FROM DOMST020_TEMP);")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("UPDATE DOSYS010_TEMP SET TRAN_COUNT = (SELECT COUNT(1) FROM DOTRN020_TEMP);")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("INSERT INTO DOSYS010_MASTER SELECT * FROM DOSYS010_TEMP;")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("TRUNCATE TABLE DOSYS010_TEMP;")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("TRUNCATE TABLE DOMST020_TEMP;")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("TRUNCATE TABLE DOTRN020_TEMP;")
                sw3.WriteLine("COMMIT;")
                sw3.WriteLine("EXIT")
                sw3.Close()

                Process.Start(Disk & ":\dump\script\create_user.bat")

                processrunning = True
                While processrunning
                    tempcount = 0
                    myProcesses = Process.GetProcesses()
                    For Each myProcess In myProcesses
                        If UCase(myProcess.ProcessName) = "CMD" Then
                            tempcount = 1
                        End If
                    Next
                    If tempcount = 0 Then
                        processrunning = False
                    End If
                    If processrunning = True Then
                        Thread.Sleep(2000)
                    End If
                End While
            Next
        End If


        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub
    Private Sub option617()   'Twenty Session Batch Job

        Dim gsolid As String
        Dim gsolset As String = UCase(InputBox("Enter SOLSET", "Enter Value", ""))
        If gsolset = "" Then
            MsgBox("No SOLSET entered.  Exiting application")
            Exit Sub
        End If

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()
        sql = "PKGLAPTOP20SESSION.GENERATE_SOL_LIST"
        Dim cmd1 As New OracleCommand(sql, conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.Add("GSOLSET", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = gsolset
        cmd1.ExecuteNonQuery()
        conn.Close()

        If Directory.Exists(Disk & ":\20Session") Then
            System.IO.Directory.Delete(Disk & ":\20Session", True)
            Directory.CreateDirectory(Disk & ":\20Session")
        Else
            Directory.CreateDirectory(Disk & ":\20Session")
        End If

        '' Executing programme for first 20 SOLs

        conn.Open()
        sql = "SELECT TD_SOLID FROM (SELECT TD_SOLID FROM C_TEMPDATA WHERE TD_PROCESSID = '20SESSION_SOL' ORDER BY TD_SOLID DESC) WHERE ROWNUM < 21"
        Dim cmd2 As New OracleCommand(sql, conn)
        Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
        dr2 = cmd2.ExecuteReader()
        While dr2.Read()
            gsolid = dr2("TD_SOLID")

            Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\20Session\" & gsolid & ".bat")
            sw2.WriteLine("@echo off")
            sw2.WriteLine("sqlplus /nolog @" & Disk & ":\20Session\" & gsolid & ".sql")
            sw2.Close()

            Dim sw3 As StreamWriter = New StreamWriter(Disk & ":\20Session\" & gsolid & ".sql")
            sw3.WriteLine("connect " & username & "/" & username & ";")
            sw3.WriteLine("EXEC PKGLAPTOP20SESSION.GENERATE('" & gsolid & "','AAA');")
            sw3.WriteLine("EXIT")
            sw3.Close()

            Process.Start(Disk & ":\20Session\" & gsolid & ".bat")

        End While
        dr2.Close()
        conn.Close()

        tempcount = 1
        Do Until tempcount = 0
            conn.Open()
            sql = "SELECT PKGLAPTOP20SESSION.TW_SESSION_PENDING_TO_START AAA FROM DUAL"
            Dim cmd4 As New OracleCommand(sql, conn)
            Dim dr4 As OracleDataReader = cmd4.ExecuteReader()
            dr4 = cmd4.ExecuteReader()
            While dr4.Read()
                tempcount = dr4("AAA")
            End While
            If tempcount > 0 Then
                sql = "SELECT TD_SOLID FROM (SELECT TD_SOLID FROM C_TEMPDATA WHERE TD_PROCESSID = '20SESSION_SOL' AND TD_TEXT1 IS NULL ORDER BY TD_SOLID DESC) WHERE ROWNUM <= PKGLAPTOP20SESSION.TWENTY_SESSION_NEXT_BATCH"
                Dim cmd3 As New OracleCommand(sql, conn)
                Dim dr3 As OracleDataReader = cmd3.ExecuteReader()
                dr3 = cmd3.ExecuteReader()
                While dr3.Read()
                    gsolid = dr3("TD_SOLID")

                    Dim sw4 As StreamWriter = New StreamWriter(Disk & ":\20Session\" & gsolid & ".bat")
                    sw4.WriteLine("@echo off")
                    sw4.WriteLine("sqlplus /nolog @" & Disk & ":\20Session\" & gsolid & ".sql")
                    sw4.Close()

                    Dim sw5 As StreamWriter = New StreamWriter(Disk & ":\20Session\" & gsolid & ".sql")
                    sw5.WriteLine("connect " & username & "/" & username & ";")
                    sw5.WriteLine("EXEC PKGLAPTOP20SESSION.GENERATE('" & gsolid & "','AAA');")
                    sw5.WriteLine("EXIT")
                    sw5.Close()

                    Process.Start(Disk & ":\20Session\" & gsolid & ".bat")

                End While
                dr3.Close()
                Thread.Sleep(3000)
            End If
            dr4.Close()
            conn.Close()
        Loop







        'kkkkkkkkkkkkkkkkk

        'tempcount = 20
        'Do Until tempcount = 0
        '    tempcount = tempcount - 1
        'Loop



        'sql = "select count(1) from acmaster where acno = '" & acno & "'"
        'cmd.Connection = cnn
        'cmd.CommandText = sql
        'dr = cmd.ExecuteReader()
        'If dr.Read = True Then
        '    recordcount = dr(0)
        'End If


        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")



        ''' creating 20Session Folder


        ''' Inserts SOL



        'Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\20Session\script\create_sol.bat")
        'sw2.WriteLine("@echo off")
        'sw2.WriteLine("sqlplus /nolog @" & Disk & ":\20Session\script\create_sol.sql")
        'sw2.Close()

        'Dim sw3 As StreamWriter = New StreamWriter(Disk & ":\20Session\script\create_sol.sql")
        'sw3.WriteLine("connect " & username & "/" & username & ";")
        'sw3.WriteLine("DELETE FROM C_TEMPDATA WHERE TD_PROCESSID = '20SESSION_SOL';")
        'sw3.WriteLine("INSERT INTO C_TEMPDATA (TD_PROCESSID,TD_SOLID) SELECT '20SESSION_SOL',SOL_ID FROM SST WHERE SET_ID = 'ALL';")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("EXIT")
        'sw3.Close()

        'Process.Start(Disk & ":\20Session\script\create_sol.bat")




        'Dim file1 As String = "c:\du\" & tempvar & "\" & tempvar & "_FRESH_INFLOW.txt"
        'If File.Exists(file1) Then

        '    File.Delete(file1)

        'End If


        'Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
        'sw0.WriteLine("@echo off")
        'sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
        'sw0.Close()

        'Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
        'sw.WriteLine("connect three/three;")
        'sw.WriteLine("drop table dosys010;")
        'sw.WriteLine("drop table domst020;")
        'sw.WriteLine("drop table dotrn020;")
        'sw.WriteLine("exit")
        'sw.Close()

        'Process.Start(Disk & ":\dump\script\create_user.bat")

        'Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
        'sw2.WriteLine("@echo off")
        'sw2.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
        'sw2.Close()

        'Dim sw3 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
        'sw3.WriteLine("connect three/three;")
        'sw3.WriteLine("INSERT INTO DOSYS010_TEMP (BRCP_BANK_CODE,BRCP_BRANCH_CODE,BRCP_BANK_NAME,BRCP_BRANCH_NAME,ADD_DATE) SELECT BRCP_BANK_CODE,BRCP_BRANCH_CODE,BRCP_BANK_NAME,BRCP_BRANCH_NAME,SYSTIMESTAMP FROM DOSYS010;")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("INSERT INTO DOMST020_TEMP SELECT ACCT_BANK_CD,ACCT_BRANCH_CD,ACCT_SUBSYS_CD,ACCT_AC_NO,ACCT_GLOBAL_ID,ACCT_NAME,ACCT_AC_OPEN_DATE,ACCT_CLEAR_BAL,ACCT_TOTAL_BAL,ACCT_STATUS FROM DOMST020 WHERE ACCT_SUBSYS_CD IN ('OSLEDL','DLEDL') AND ACCT_STATUS NOT IN (98) AND ACCT_AC_CLOSED_DATE IS NULL;")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("INSERT INTO DOMST020_MASTER SELECT * FROM DOMST020_TEMP;")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("INSERT INTO DOTRN020_TEMP SELECT MTRN_BANK_CD,MTRN_BRANCH_CD,MTRN_DATE,MTRN_SCROLL_NO_X,MTRN_SCROLL_NO_9,MTRN_VOUCHER_SRNO,MTRN_SUBSYS_CD,MTRN_AC_NO,MTRN_ASOF_DATE,MTRN_INS_TYPE,MTRN_AMOUNT,MTRN_PARTICULARS,MTRN_STATUS FROM DOTRN020 WHERE (MTRN_SUBSYS_CD,MTRN_AC_NO) IN (SELECT ACCT_SUBSYS_CD,ACCT_AC_NO FROM DOMST020_TEMP);")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("INSERT INTO DOTRN020_MASTER SELECT * FROM DOTRN020_TEMP;")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("UPDATE DOSYS010_TEMP SET MAX_TRAN_DATE = (SELECT MAX(MTRN_DATE) FROM DOTRN020);")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("UPDATE DOSYS010_TEMP SET AC_COUNT = (SELECT COUNT(1) FROM DOMST020_TEMP);")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("UPDATE DOSYS010_TEMP SET TRAN_COUNT = (SELECT COUNT(1) FROM DOTRN020_TEMP);")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("INSERT INTO DOSYS010_MASTER SELECT * FROM DOSYS010_TEMP;")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("TRUNCATE TABLE DOSYS010_TEMP;")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("TRUNCATE TABLE DOMST020_TEMP;")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("TRUNCATE TABLE DOTRN020_TEMP;")
        'sw3.WriteLine("COMMIT;")
        'sw3.WriteLine("EXIT")
        'sw3.Close()

        'Process.Start(Disk & ":\dump\script\create_user.bat")


        'Dim dumppath As String = InputBox("Enter path in which is backups are placed", "Enter Value", "D:\1")
        'If dumppath = "" Then
        '    MsgBox("No data entered.  Exiting application")
        '    Exit Sub
        'End If

        'Dim tempcount As Integer
        'Dim processrunning As Boolean
        'Dim myProcesses() As Process
        'Dim myProcess As Process

        'Dim dirs As String() = Directory.GetFiles(dumppath, "*.dmp")
        'Dim dir As String
        'Dim totalfiles As Integer
        'Dim filename As String
        'Dim uname As String

        'totalfiles = dirs.Length

        'If totalfiles = 0 Then
        '    MsgBox("No .dmp exists in the folder " & dumppath, MsgBoxStyle.Critical, "Error")
        'Else
        '    For Each dir In dirs
        '        tempcount = tempcount + 1
        '        filename = GetFileName(dir)
        '        uname = "three"

        '        Dim sw0 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
        '        sw0.WriteLine("@echo off")
        '        sw0.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
        '        sw0.Close()

        '        Dim sw As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
        '        sw.WriteLine("connect three/three;")
        '        sw.WriteLine("drop table dosys010;")
        '        sw.WriteLine("drop table domst020;")
        '        sw.WriteLine("drop table dotrn020;")
        '        sw.WriteLine("exit")
        '        sw.Close()

        '        Process.Start(Disk & ":\dump\script\create_user.bat")

        '        processrunning = True
        '        While processrunning
        '            tempcount = 0
        '            myProcesses = Process.GetProcesses()
        '            For Each myProcess In myProcesses
        '                If UCase(myProcess.ProcessName) = "CMD" Then
        '                    tempcount = 1
        '                End If
        '            Next
        '            If tempcount = 0 Then
        '                processrunning = False
        '            End If
        '            If processrunning = True Then
        '                Thread.Sleep(2000)
        '            End If
        '        End While

        '        Dim sw1 As StreamWriter = New StreamWriter(Disk & ":\dump\script\import_user_br.bat")
        '        sw1.WriteLine("@echo off")
        '        sw1.WriteLine("imp " & uname & "/" & uname & "@ten file=" & dir & " TABLES=(domst020, dosys010, dotrn020) GRANTS=N CONSTRAINTS=N INDEXES=N IGNORE=Y")
        '        sw1.Close()
        '        Process.Start(Disk & ":\dump\script\import_user_br.bat")

        '        processrunning = True
        '        While processrunning
        '            tempcount = 0
        '            myProcesses = Process.GetProcesses()
        '            For Each myProcess In myProcesses
        '                If UCase(myProcess.ProcessName) = "CMD" Then
        '                    tempcount = 1
        '                End If
        '            Next
        '            If tempcount = 0 Then
        '                processrunning = False
        '            End If
        '            If processrunning = True Then
        '                Thread.Sleep(2000)
        '            End If
        '        End While

        '        Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.bat")
        '        sw2.WriteLine("@echo off")
        '        sw2.WriteLine("sqlplus /nolog @" & Disk & ":\dump\script\create_user.sql")
        '        sw2.Close()

        '        Dim sw3 As StreamWriter = New StreamWriter(Disk & ":\dump\script\create_user.sql")
        '        sw3.WriteLine("connect three/three;")
        '        sw3.WriteLine("INSERT INTO DOSYS010_TEMP (BRCP_BANK_CODE,BRCP_BRANCH_CODE,BRCP_BANK_NAME,BRCP_BRANCH_NAME,ADD_DATE) SELECT BRCP_BANK_CODE,BRCP_BRANCH_CODE,BRCP_BANK_NAME,BRCP_BRANCH_NAME,SYSTIMESTAMP FROM DOSYS010;")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("INSERT INTO DOMST020_TEMP SELECT ACCT_BANK_CD,ACCT_BRANCH_CD,ACCT_SUBSYS_CD,ACCT_AC_NO,ACCT_GLOBAL_ID,ACCT_NAME,ACCT_AC_OPEN_DATE,ACCT_CLEAR_BAL,ACCT_TOTAL_BAL,ACCT_STATUS FROM DOMST020 WHERE ACCT_SUBSYS_CD IN ('OSLEDL','DLEDL') AND ACCT_STATUS NOT IN (98) AND ACCT_AC_CLOSED_DATE IS NULL;")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("INSERT INTO DOMST020_MASTER SELECT * FROM DOMST020_TEMP;")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("INSERT INTO DOTRN020_TEMP SELECT MTRN_BANK_CD,MTRN_BRANCH_CD,MTRN_DATE,MTRN_SCROLL_NO_X,MTRN_SCROLL_NO_9,MTRN_VOUCHER_SRNO,MTRN_SUBSYS_CD,MTRN_AC_NO,MTRN_ASOF_DATE,MTRN_INS_TYPE,MTRN_AMOUNT,MTRN_PARTICULARS,MTRN_STATUS FROM DOTRN020 WHERE (MTRN_SUBSYS_CD,MTRN_AC_NO) IN (SELECT ACCT_SUBSYS_CD,ACCT_AC_NO FROM DOMST020_TEMP);")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("INSERT INTO DOTRN020_MASTER SELECT * FROM DOTRN020_TEMP;")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("UPDATE DOSYS010_TEMP SET MAX_TRAN_DATE = (SELECT MAX(MTRN_DATE) FROM DOTRN020);")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("UPDATE DOSYS010_TEMP SET AC_COUNT = (SELECT COUNT(1) FROM DOMST020_TEMP);")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("UPDATE DOSYS010_TEMP SET TRAN_COUNT = (SELECT COUNT(1) FROM DOTRN020_TEMP);")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("INSERT INTO DOSYS010_MASTER SELECT * FROM DOSYS010_TEMP;")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("TRUNCATE TABLE DOSYS010_TEMP;")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("TRUNCATE TABLE DOMST020_TEMP;")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("TRUNCATE TABLE DOTRN020_TEMP;")
        '        sw3.WriteLine("COMMIT;")
        '        sw3.WriteLine("EXIT")
        '        sw3.Close()

        '        Process.Start(Disk & ":\dump\script\create_user.bat")

        '        processrunning = True
        '        While processrunning
        '            tempcount = 0
        '            myProcesses = Process.GetProcesses()
        '            For Each myProcess In myProcesses
        '                If UCase(myProcess.ProcessName) = "CMD" Then
        '                    tempcount = 1
        '                End If
        '            Next
        '            If tempcount = 0 Then
        '                processrunning = False
        '            End If
        '            If processrunning = True Then
        '                Thread.Sleep(2000)
        '            End If
        '        End While
        '    Next
        'End If


        'MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Private Sub option44()   'Execute query and generate multiple files

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim SOLID As String
        Dim SOLNAME As String
        Dim FSOLID As String
        Dim TEMPCOUNT As Integer = 0

        processmessage("Fetching SOLID")
        sql = "SELECT BRCP_BRANCH_CODE,BRCP_BRANCH_NAME,FSOLID FROM DOSYS010_MASTER"
        Dim cmd2 As New OracleCommand(sql, conn)
        Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
        While dr2.Read
            TEMPCOUNT = TEMPCOUNT + 1
            processmessage("Branch No - " & TEMPCOUNT)
            SOLID = dr2.Item("BRCP_BRANCH_CODE").ToString.Trim
            SOLNAME = dr2.Item("BRCP_BRANCH_NAME").ToString.Trim
            FSOLID = dr2.Item("FSOLID").ToString.Trim

            sql = "PKGMISTOOL.TEMP"
            Dim cmd4 As New OracleCommand(sql, conn)
            cmd4.CommandType = CommandType.StoredProcedure
            cmd4.Parameters.Add("SOLID", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = SOLID
            cmd4.ExecuteNonQuery()

            sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY TO_NUMBER(SOLID),SERIALNO,SUBSERIALNO"
            display_in_File(sql, "C:\du\" & SOLID & "_" & FSOLID & "_" & SOLNAME & ".txt")

        End While
        dr2.Close()

        MsgBox("Process completed successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Sub option45()      'PMJDY Campaign

        ' Checking whether BACOPEN,SB.TXT files exists

        processmessage("Checking files")

        file1 = "c:\du\BACOPEN.txt"
        file2 = "c:\du\SB.txt"

        checkfile(file1, "Rename the file 40998_XX-XX-XXXX_AC1.TXT as BACEOPEN.TXT and place in c:/du folder")
        checkfile(file2, "Rename the file from tabdata as SB.TXT and place in c:/du folder")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - PKGEMAIL113.DATAID_1134")

        sql = "PKGEMAIL113.DATAID_1134"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd4.ExecuteNonQuery()

        sendemail("smgbmis@gmail.com", "ten", username, username)

    End Sub

    Sub option47()      'Business Figures As On 30-09-2014

        'Checking whether BACOPEN,SB.TXT files exists

        'processmessage("Checking files")

        'file1 = "c:\du\9106_16908570.rpt"
        'file2 = "c:\du\SB.txt"

        'checkfile(file1, "Place the file in c:/du folder")
        'checkfile(file2, "Rename the file from tabdata as SB.TXT and place in c:/du folder")

        'uploadfiledata(file1, username, "Y")
        'uploadfiledata(file2, username, "N")

        'Delete existing data, if any, from c_du table

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

        End If

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        'processmessage("Package - Data ID - 1171")       'KGB Day Book

        'sql = "PKGEMAIL117.DATAID_1171"
        'Dim cmd6 As New OracleCommand(sql, conn)
        'cmd6.CommandType = CommandType.StoredProcedure
        'cmd6.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        'cmd6.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1172")       'Cash Balance and Bankers Account

        'sql = "PKGEMAIL117.DATAID_1172"
        'Dim cmd6 As New OracleCommand(sql, conn)
        'cmd6.CommandType = CommandType.StoredProcedure
        'cmd6.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        'cmd6.ExecuteNonQuery()

        processmessage("Package - Data ID - 1175")       'NPA In Out

        sql = "PKGEMAIL117.DATAID_1175"
        Dim cmd6 As New OracleCommand(sql, conn)
        cmd6.CommandType = CommandType.StoredProcedure
        cmd6.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd6.ExecuteNonQuery()

        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'NPAINOUT' ORDER BY SOLID"
        display_in_File(sql, "C:\du\SMS_" & RptDate.ToString("ddMMyyyy") & ".txt")
        Process.Start("C:\du\SMS_" & RptDate.ToString("ddMMyyyy") & ".txt")

        sendemail("smgbmis@gmail.com", "ten", username, username)
        'sendemail("dipsdot@gmail.com", "ten", username, username)

    End Sub

    Sub option58()      'BOD Mails

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        file1 = "c:\du\friday.txt"
        checkfile(file1, "Place last friday file in C:/DU folder as 'FRIDAY.txt'")

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

        End If

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - Data ID - 1171")       'KGB Day Book

        sql = "PKGEMAIL117.DATAID_1171"
        Dim cmd6 As New OracleCommand(sql, conn)
        cmd6.CommandType = CommandType.StoredProcedure
        cmd6.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd6.ExecuteNonQuery()

        'sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'TEMPDATA' ORDER BY SERIALNO,SUBSERIALNO"
        'display_in_File(sql, "C:\du\C_TEMPDATA_ADD.txt")
        'Process.Start("C:\du\C_TEMPDATA_ADD.txt")

        processmessage("Package - Data ID - 1172")       'Cash Balance and Bankers Account

        sql = "PKGEMAIL117.DATAID_1172"
        Dim cmd7 As New OracleCommand(sql, conn)
        cmd7.CommandType = CommandType.StoredProcedure
        cmd7.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd7.ExecuteNonQuery()

        'processmessage("Package - Data ID - 1174")       'Balance in Migration Pooling A/c

        'sql = "PKGEMAIL117.DATAID_1174"
        'Dim cmd9 As New OracleCommand(sql, conn)
        'cmd9.CommandType = CommandType.StoredProcedure
        'cmd9.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        'cmd9.ExecuteNonQuery()

        processmessage("Package - Data ID - 1173")       'KGB Business Progress Report

        sql = "PKGEMAIL117.DATAID_1173"
        Dim cmd8 As New OracleCommand(sql, conn)
        cmd8.CommandType = CommandType.StoredProcedure
        cmd8.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd8.ExecuteNonQuery()

        'sendemail("smgbmis2@gmail.com", "ten", username, username)
        sendemail("mu@kgbmis.in", "ten", username, username)
        'sendemail("kgbmis1@gmail.com", "ten", username, username)

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        processmessage("Package - Data ID - 1193")       'NRE Transaction CASA

        sql = "PKGEMAIL119.DATAID_1193"
        Dim cmd10 As New OracleCommand(sql, conn)
        cmd10.CommandType = CommandType.StoredProcedure
        cmd10.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
        cmd10.ExecuteNonQuery()

        'sendemail("kgbmis1@gmail.com", "ten", username, username)
        sendemail("br@kgbmis.in", "ten", username, username)
        'sendemail("dipsdot@gmail.com", "ten", username, username)

    End Sub

    Sub option46()      'SMS File Creation

        ' Creating SMS in the following File Format 
        '        Bulk SMS data required.
        '        i.Country(code)
        '        ii.Mobile(number)
        '        iii.Customer(ID)
        '        iv.Account(number)
        '        v.Message(Literal)

        'Format()
        'xx|xxxxxxxxxx|||xxxxxxxxxxxxxxxxxxxxxx

        'Example.
        '91|9447954691|||Test message

        Dim msg As String
        Dim desig As String
        Dim office As String = ""
        Dim department As String = ""
        Dim Filename As String = "C:\temp\smsupd.txt"
        Dim Include_branch As String = ""
        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        desig = InputBox("Enter designation (ALL,GM,RM,SM,MG,CH)", "Enter Value", "ALL")
        office = InputBox("Enter Office type(ALL,RO,HO,BR,HD)", "Enter Value", "ALL")
        department = InputBox("Enter Department(CW,CS,HW,IT,RL,PD)", "Enter Value", "ALL")

        If office.ToUpper() = "ALL" Or office.ToUpper() = "BR" Then

            If desig.ToUpper() <> "GM" And desig.ToUpper() <> "RM" And desig.ToUpper() <> "CH" And department.ToUpper() = "ALL" Then

                Include_branch = "Y"
            End If

        Else
            Include_branch = "N"
        End If

        msg = InputBox("Enter the message")

        If desig.ToUpper = "ALL" And office.ToUpper = "ALL" And department.ToUpper = "ALL" Then
            sql = "SELECT MOBILE_NUM,DESIGNATION FROM Z_CUG"

        ElseIf desig.ToUpper = "ALL" And office.ToUpper = "ALL" Then

            department = "'" + department + "'"
            department = department.Replace(",", "','")
            department = department.ToUpper()

            sql = "SELECT MOBILE_NUM,DESIGNATION FROM Z_CUG WHERE DEPARTMENT IN (" + department + ")"

        ElseIf desig.ToUpper = "ALL" And department.ToUpper = "ALL" Then
            office = "'" + office + "'"
            office = office.Replace(",", "','")
            office = office.ToUpper()

            sql = "SELECT MOBILE_NUM,DESIGNATION FROM Z_CUG WHERE OFFICE_TYPE IN (" + office + ")"

        ElseIf office.ToUpper = "ALL" And department.ToUpper = "ALL" Then

            desig = "'" + desig + "'"
            desig = desig.Replace(",", "','")
            desig = desig.ToUpper()

            sql = "SELECT MOBILE_NUM,DESIGNATION FROM Z_CUG WHERE DESIGNATION IN (" + desig + ")"

        ElseIf desig.ToUpper = "ALL" Then

            department = "'" + department + "'"
            department = department.Replace(",", "','")
            department = department.ToUpper()

            office = "'" + office + "'"
            office = office.Replace(",", "','")
            office = office.ToUpper()
            sql = "SELECT MOBILE_NUM,DESIGNATION FROM Z_CUG WHERE DEPARTMENT IN (" + department + ")  AND OFFICE_TYPE IN (" + office + ")"

        ElseIf department.ToUpper = "ALL" Then

            office = "'" + office + "'"
            office = office.Replace(",", "','")
            office = office.ToUpper()

            desig = "'" + desig + "'"
            desig = desig.Replace(",", "','")
            desig = desig.ToUpper()

            sql = "SELECT MOBILE_NUM,DESIGNATION FROM Z_CUG WHERE  DESIGNATION IN (" + desig + ") AND OFFICE_TYPE IN (" + office + ")"

        ElseIf office.ToUpper = "ALL" Then

            desig = "'" + desig + "'"
            desig = desig.Replace(",", "','")
            desig = desig.ToUpper()

            department = "'" + department + "'"
            department = department.Replace(",", "','")
            department = department.ToUpper()
            sql = "SELECT MOBILE_NUM,DESIGNATION FROM Z_CUG WHERE DEPARTMENT IN (" + department + ") AND DESIGNATION IN (" + desig + ") "

        Else

            department = "'" + department + "'"
            department = department.Replace(",", "','")
            department = department.ToUpper()

            office = "'" + office + "'"
            office = office.Replace(",", "','")
            office = office.ToUpper()

            desig = "'" + desig + "'"
            desig = desig.Replace(",", "','")
            desig = desig.ToUpper()
            sql = "SELECT MOBILE_NUM,DESIGNATION FROM Z_CUG WHERE DEPARTMENT IN (" + department + ") AND DESIGNATION IN (" + desig + ") AND OFFICE_TYPE IN (" + office + ")"

        End If

        Dim cmd1 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd1.ExecuteReader()
        Dim linedata As String
        Dim sw As StreamWriter = New StreamWriter(Filename.ToString())
        While dr.Read()
            linedata = "91|"
            linedata = linedata + dr("MOBILE_NUM").ToString() + "|||" + msg
            sw.WriteLine(linedata)
        End While
        dr.Close()
        sw.Close()

        If Include_branch = "Y" Then

            sql = "SELECT SUBSTR(SOLID2,3,3) SOL FROM C_MISONLINEDATE WHERE SOLID2 > 40101"

            Dim cmd2 As New OracleCommand(sql, conn)
            Dim dr1 As OracleDataReader = cmd2.ExecuteReader()

            Dim sw1 As StreamWriter = New StreamWriter(Filename.ToString(), True)
            While dr1.Read()
                linedata = "91|9400999"
                linedata = linedata + dr1("SOL").ToString() + "|||" + msg
                sw1.WriteLine(linedata)
            End While
            dr1.Close()
            sw1.Close()

        End If
        conn.Close()
        MsgBox("File generated Successfully in C:\temp folder")
    End Sub

    Private Sub releaseEXCELObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Sub option55()      'NPA Threat For Next 7 Days - Email Generation

        'Dim dirs As String() = Directory.GetFiles("c:\du")
        'Dim dir As String
        'Dim totalfiles As Integer
        'Dim tempcount As Integer = 0

        'totalfiles = dirs.Length

        'If totalfiles = 0 Then

        '    processmessage("")
        '    MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        'Else

        '    For Each dir In dirs

        '        tempcount = tempcount + 1

        '        If tempcount = 1 Then

        '            uploadfiledata_without_trim(dir, username, "Y")

        '        Else

        '            uploadfiledata_without_trim(dir, username, "N")

        '        End If

        '    Next

        'End If

        'processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - PKGEMAIL115.DATAID_1155_CREATE_MAIL")

        sql = "PKGEMAIL115.DATAID_1155_CREATE_MAIL"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Package - PKGEMAIL116.DATAID_1163")

        sql = "PKGEMAIL116.DATAID_1163"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd5.ExecuteNonQuery()

        'sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'NPAIMPACT' ORDER BY SOLID"
        'display_in_File(sql, "C:\du\SMS_NPAIMPACT.txt")
        'Process.Start("C:\du\SMS_NPAIMPACT.txt")

        'sendemail_pnpa("smgbmis2@gmail.com", "ten", username, username)
        sendemail_pnpa("nt@kgbmis.in", "ten", username, username)

    End Sub

    Sub option57()      'Predefined Day End Check Validation

        'processmessage("Checking files")

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")
            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

        End If

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - PKGEMAIL115.PNPA_7DAY_PROCESS_DATA")

        sql = "PKGEMAIL115.PNPA_7DAY_PROCESS_DATA"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Package - PKGEMAIL116.DATAID_1164")

        sql = "PKGEMAIL116.DATAID_1164"
        Dim cmd5 As New OracleCommand(sql, conn)
        cmd5.CommandType = CommandType.StoredProcedure
        cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd5.ExecuteNonQuery()

        processmessage("Package - PKGEMAIL116.DATAID_1165")

        sql = "PKGEMAIL116.DATAID_1165"
        Dim cmd6 As New OracleCommand(sql, conn)
        cmd6.CommandType = CommandType.StoredProcedure
        cmd6.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd6.ExecuteNonQuery()

        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'DEC' ORDER BY SOLID"
        display_in_File(sql, "C:\du\SMS_DEC.txt")

        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'DEC1' ORDER BY SOLID"
        display_in_File(sql, "C:\du\SMS_DEC_1.txt")

        Process.Start("C:\du\SMS_DEC_1.txt")
        Process.Start("C:\du\SMS_DEC.txt")

        'sendemail("smgbmis3@gmail.com", "ten", username, username)
        sendemail("mu@kgbmis.in", "ten", username, username)
        'sendemail("pdec@kgbmis.in", "ten", username, username)

    End Sub

    Private Sub Option54()              'NPA Threat For Next 7 Days - Excel Creation

        If Process.GetProcessesByName("excel").GetLength(0) > 0 Then

            MessageBox.Show("Please close Excel Windows if any and click OK")
            ' Exit Sub

        End If

        Dim yn = UCase(InputBox("Reload data?", "Confirm reload", "Y"))

        If yn = "Y" Then

            Dim dirs As String() = Directory.GetFiles("c:\du")
            Dim dir As String
            Dim totalfiles As Integer
            Dim tempcount As Integer = 0

            totalfiles = dirs.Length

            If totalfiles = 0 Then

                processmessage("")
                MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

            Else

                For Each dir In dirs

                    tempcount = tempcount + 1

                    If tempcount = 1 Then

                        uploadfiledata_without_trim(dir, username, "Y")

                    Else

                        uploadfiledata_without_trim(dir, username, "N")

                    End If

                Next

            End If

            processmessage("Deleting existing data")

            oracle_execute_non_query("ten", username, username, "truncate table z_email")

            ' Calling packages

            Dim sql As String
            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - PKGEMAIL115.PNPA_7DAY_PROCESS_DATA")

            sql = "PKGEMAIL115.PNPA_7DAY_PROCESS_DATA"
            Dim cmd4 As New OracleCommand(sql, conn)
            cmd4.CommandType = CommandType.StoredProcedure
            cmd4.ExecuteNonQuery()

            processmessage("Package - PKGEMAIL116.DATAID_1163")

            sql = "PKGEMAIL116.DATAID_1163"
            Dim cmd5 As New OracleCommand(sql, conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd5.ExecuteNonQuery()

            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'NPAIMPACT' ORDER BY SOLID"
            display_in_File(sql, "C:\du\SMS_NPAIMPACT.txt")
            Process.Start("C:\du\SMS_NPAIMPACT.txt")

            conn.Close()

        End If

        '===========================================================

        Dim oracle_cnn_string As String = "Data Source=ten; User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        processmessage("Initializing ...")

        oracle_execute_non_query("ten", username, username, "UPDATE C_MISADV SET NUMBER20 = (CASE WHEN NPAMAIN = 'M' THEN 1 WHEN NPAMAIN = 'S' THEN 2 WHEN NPAMAIN = 'C' THEN 3 WHEN NPAMAIN = 'N' THEN 4 WHEN NPAMAIN = 'F' THEN 5 ELSE 0 END)")
        oracle_execute_non_query("ten", username, username, "update c_misadv A SET (TEXT11,TEXT12,TEXT13,DATE11,DATE12,NUMBER11,DATE13,TEXT14,NUMBER12,NUMBER13,NUMBER14,NUMBER15,NUMBER16,NUMBER17,NUMBER18) = (select TEXT1,TEXT2,TEXT3,DATE1,DATE2,NUMBER1,DATE3,TEXT4,NUMBER2,NUMBER3,NUMBER4,NUMBER5,NUMBER6,NUMBER7,NUMBER8 FROM C_MISADV B WHERE B.NPAMAIN='Z' AND A.SOLID = B.SOLID) WHERE A.NPAMAIN='M'")
        oracle_execute_non_query("ten", username, username, "update c_misadv A SET (NUMBER21, DATE20) = (select NUMBER2,DATE3 FROM C_MISADV B WHERE B.NPAMAIN='M' AND A.ACNO = B.ACNO) WHERE A.NPAMAIN IN ('M','S','C','N','F')")
        oracle_execute_non_query("ten", username, username, "COMMIT")

        oracle_conn.Close()

        Dim SOL_SET As String
        Dim START_AMNT As Double
        Dim END_AMNT As Double
        Dim FILE_NAME As String


        Call GEN_NPA_XL("ROKSD", 500000, 99999999999.99, "ROKSD")
        Call GEN_NPA_XL("ROKSD", 300000, 499999.99, "ROKSD")
        Call GEN_NPA_XL("ROKSD", 100000, 299999.99, "ROKSD")
        Call GEN_NPA_XL("ROKSD", -99999999999.99, 99999.99, "ROKSD")


        Call GEN_NPA_XL("ROKNR", 500000, 99999999999.99, "ROKNR")
        Call GEN_NPA_XL("ROKNR", 300000, 499999.99, "ROKNR")
        Call GEN_NPA_XL("ROKNR", 100000, 299999.99, "ROKNR")
        Call GEN_NPA_XL("ROKNR", -99999999999.99, 99999.99, "ROKNR")


        Call GEN_NPA_XL("ROTLY", 500000, 99999999999.99, "ROTLY")
        Call GEN_NPA_XL("ROTLY", 300000, 499999.99, "ROTLY")
        Call GEN_NPA_XL("ROTLY", 100000, 299999.99, "ROTLY")
        Call GEN_NPA_XL("ROTLY", -99999999999.99, 99999.99, "ROTLY")


        Call GEN_NPA_XL("ROKPT", 500000, 99999999999.99, "ROKPT")
        Call GEN_NPA_XL("ROKPT", 300000, 499999.99, "ROKPT")
        Call GEN_NPA_XL("ROKPT", 100000, 299999.99, "ROKPT")
        Call GEN_NPA_XL("ROKPT", -99999999999.99, 99999.99, "ROKPT")


        Call GEN_NPA_XL("ROTSR", 500000, 99999999999.99, "ROTSR")
        Call GEN_NPA_XL("ROTSR", 300000, 499999.99, "ROTSR")
        Call GEN_NPA_XL("ROTSR", 100000, 299999.99, "ROTSR")
        Call GEN_NPA_XL("ROTSR", -99999999999.99, 99999.99, "ROTSR")


        Call GEN_NPA_XL("ROKKD", 500000, 99999999999.99, "ROKKD")
        Call GEN_NPA_XL("ROKKD", 300000, 499999.99, "ROKKD")
        Call GEN_NPA_XL("ROKKD", 100000, 299999.99, "ROKKD")
        Call GEN_NPA_XL("ROKKD", -99999999999.99, 99999.99, "ROKKD")


        Call GEN_NPA_XL("ROMPM", 500000, 99999999999.99, "ROMPM")
        Call GEN_NPA_XL("ROMPM", 300000, 499999.99, "ROMPM")
        Call GEN_NPA_XL("ROMPM", 100000, 299999.99, "ROMPM")
        Call GEN_NPA_XL("ROMPM", -99999999999.99, 99999.99, "ROMPM")


        Call GEN_NPA_XL("ROEKM", 500000, 99999999999.99, "ROEKM")
        Call GEN_NPA_XL("ROEKM", 300000, 499999.99, "ROEKM")
        Call GEN_NPA_XL("ROEKM", 100000, 299999.99, "ROEKM")
        Call GEN_NPA_XL("ROEKM", -99999999999.99, 99999.99, "ROEKM")


        Call GEN_NPA_XL("ROKTM", 500000, 99999999999.99, "ROKTM")
        Call GEN_NPA_XL("ROKTM", 300000, 499999.99, "ROKTM")
        Call GEN_NPA_XL("ROKTM", 100000, 299999.99, "ROKTM")
        Call GEN_NPA_XL("ROKTM", -99999999999.99, 99999.99, "ROKTM")


        Call GEN_NPA_XL("ROTVM", 500000, 99999999999.99, "ROTVM")
        Call GEN_NPA_XL("ROTVM", 300000, 499999.99, "ROTVM")
        Call GEN_NPA_XL("ROTVM", 100000, 299999.99, "ROTVM")
        Call GEN_NPA_XL("ROTVM", -99999999999.99, 99999.99, "ROTVM")


        Call GEN_NPA_XL("ALL", 300000, 99999999999.99, "RL")   'R&L
        Call GEN_NPA_XL("ALL", 100000, 299999.99, "CRM")    'CRM
        Call GEN_NPA_XL("ALL", 300000, 99999999999.99, "GM")  'GM
        Call GEN_NPA_XL("ALL", 500000, 99999999999.99, "CM")  'CHAIRMAN

        'Call GEN_NPA_XL("ALL", -9999999, 9999999999, "ALL")
        'MessageBox.Show("File generation over.")

        MsgBox("Excel file generation completed.", MsgBoxStyle.Information, "Process Completed")

        '=====================================================
        ''To ROs
        'NPA_THREAT_EXCEL_EMAIL("ROTVM", 500000, 99999999999.99, "aotvpm@yahoo.com", "ROTVM@keralagbank.com;kgb399@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTVM_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTVM_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTVM_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTVM_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROTVM)")
        'NPA_THREAT_EXCEL_EMAIL("ROEKM", 500000, 99999999999.99, "roekm.kgb@gmail.com", "ROEKM@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROEKM_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROEKM_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROEKM_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROEKM_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROEKM)")
        'NPA_THREAT_EXCEL_EMAIL("ROTSR", 500000, 99999999999.99, "smgbaotri@yahoo.co.in", "kgb400@keralagbank.com;ROKZD@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTSR_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTSR_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTSR_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTSR_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROTSR)")
        'NPA_THREAT_EXCEL_EMAIL("ROMPM", 500000, 99999999999.99, "kgbroperintalmanna@gmail.com", "ROPMA@keralagbank.com;kgb397@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROMPM_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROMPM_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROMPM_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROMPM_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROMPM)")
        'NPA_THREAT_EXCEL_EMAIL("ROKKD", 500000, 99999999999.99, "kgbroclt@gmail.com", "ROKZD@keralagbank.com;kgb393@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKKD_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKKD_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKKD_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKKD_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKKD)")
        'NPA_THREAT_EXCEL_EMAIL("ROTLY", 500000, 99999999999.99, "kgbrotly@gmail.com", "ROTLY@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTLY_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTLY_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTLY_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTLY_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROTLY)")
        'NPA_THREAT_EXCEL_EMAIL("ROKSD", 500000, 99999999999.99, "kgbroksd@gmail.com", "ROKSD@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKSD_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKSD_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKSD_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKSD_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKSD)")
        'NPA_THREAT_EXCEL_EMAIL("ROKPT", 500000, 99999999999.99, "rokpta@gmail.com", "ROKPT@keralagbank.com;kgb393@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKPT_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKPT_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKPT_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKPT_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKPT)")
        'NPA_THREAT_EXCEL_EMAIL("ROKTM", 500000, 99999999999.99, "roktm.kgb@gmail.com", "kgb660@keralagbank.com;ROKTM@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKTM_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKTM_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKTM_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKTM_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKTM)")
        'NPA_THREAT_EXCEL_EMAIL("ROKNR", 500000, 99999999999.99, "nmgbknrao@gmail.com", "ROKNR@keralagbank.com;kgb394@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKNR_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKNR_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKNR_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKNR_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKNR)")

        ''To GMs
        'NPA_THREAT_EXCEL_EMAIL("ALL", 300000, 99999999999.99, "thangavelu.ppt@gmail.com;nkkrishnankutty46876@gmail.com;haridasanv@gmail.com;srnair32474@gmail.com", "", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_GM_300000_99999999999.99.xlsx", "0", "0", "0", "NPA Threat for the next 7 Days (SOLSET : ALL; Outstanding from 300000.00 to 9999999999.99)")

        ''To CRM
        'NPA_THREAT_EXCEL_EMAIL("ALL", 100000, 299999.99, "crmwing.kgb@gmail.com", "", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_CRM_100000_299999.99.xlsx", "0", "0", "0", "NPA Threat for the next 7 Days (SOLSET : ALL; Outstanding from 100000.00 to 299999.99)")

        ''To RL
        'NPA_THREAT_EXCEL_EMAIL("ALL", 300000, 99999999999.99, "smgbrl93@gmail.com", "", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_RL_300000_99999999999.99.xlsx", "0", "0", "0", "NPA Threat for the next 7 Days (SOLSET : ALL; Outstanding from 300000.00 to 9999999999.99)")

        ''To CHM
        'NPA_THREAT_EXCEL_EMAIL("ALL", 500000, 99999999999.99, "chairmankeralagb@gmail.com", "franklinkf@gmail.com;sureshsmgb1@gmail.com;udayakumarcv@gmail.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_CM_500000_99999999999.99.xlsx", "0", "0", "0", "NPA Threat for the next 7 Days (SOLSET : ALL; Outstanding from 500000.00 to 9999999999.99)")

        'MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")
        '=====================================================

    End Sub

    Private Sub Option56()              'NPA Threat For Next 7 Days - Excel Creation Using Macro

        If Process.GetProcessesByName("excel").GetLength(0) > 0 Then

            MessageBox.Show("Please close Excel Windows if any and click OK")
            ' Exit Sub

        End If

        Dim yn = UCase(InputBox("Reload data?", "Confirm reload", "Y"))

        If yn = "Y" Then

            Dim dirs As String() = Directory.GetFiles("c:\du")
            Dim dir As String
            Dim totalfiles As Integer
            Dim tempcount As Integer = 0

            totalfiles = dirs.Length

            If totalfiles = 0 Then

                processmessage("")
                MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

            Else

                For Each dir In dirs

                    tempcount = tempcount + 1

                    If tempcount = 1 Then

                        uploadfiledata_without_trim(dir, username, "Y")

                    Else

                        uploadfiledata_without_trim(dir, username, "N")

                    End If

                Next

            End If

            processmessage("Deleting existing data")

            oracle_execute_non_query("ten", username, username, "truncate table z_email")

            ' Calling packages

            Dim sql As String
            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - PKGEMAIL115.PNPA_7DAY_PROCESS_DATA")

            sql = "PKGEMAIL115.PNPA_7DAY_PROCESS_DATA"
            Dim cmd4 As New OracleCommand(sql, conn)
            cmd4.CommandType = CommandType.StoredProcedure
            cmd4.ExecuteNonQuery()

            processmessage("Package - PKGEMAIL116.DATAID_1163")

            sql = "PKGEMAIL116.DATAID_1163"
            Dim cmd5 As New OracleCommand(sql, conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd5.ExecuteNonQuery()

            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'NPAIMPACT' ORDER BY SOLID"
            display_in_File(sql, "C:\du\SMS_NPAIMPACT.txt")

            Process.Start("C:\du\SMS_NPAIMPACT.txt")
            conn.Close()

        End If

        '===========================================================

        Dim ORACLE_CNN_STRING As String = "DATA SOURCE=TEN; USER ID= " & username & ";PASSWORD= " & username & ";"
        Dim ORACLE_CONN As New OracleConnection(ORACLE_CNN_STRING)
        ORACLE_CONN.Open()

        processmessage("INITIALIZING ...")

        oracle_execute_non_query("TEN", username, username, "UPDATE C_MISADV SET NUMBER20 = (CASE WHEN NPAMAIN = 'M' THEN 1 WHEN NPAMAIN = 'S' THEN 2 WHEN NPAMAIN = 'C' THEN 3 WHEN NPAMAIN = 'N' THEN 4 WHEN NPAMAIN = 'F' THEN 5 ELSE 0 END)")
        oracle_execute_non_query("TEN", username, username, "UPDATE C_MISADV A SET (TEXT11,TEXT12,TEXT13,DATE11,DATE12,NUMBER11,DATE13,TEXT14,NUMBER12,NUMBER13,NUMBER14,NUMBER15,NUMBER16,NUMBER17,NUMBER18) = (SELECT TEXT1,TEXT2,TEXT3,DATE1,DATE2,NUMBER1,DATE3,TEXT4,NUMBER2,NUMBER3,NUMBER4,NUMBER5,NUMBER6,NUMBER7,NUMBER8 FROM C_MISADV B WHERE B.NPAMAIN='Z' AND A.SOLID = B.SOLID) WHERE A.NPAMAIN='M'")
        oracle_execute_non_query("TEN", username, username, "UPDATE C_MISADV A SET (NUMBER21, DATE20) = (SELECT NUMBER2,DATE3 FROM C_MISADV B WHERE B.NPAMAIN='M' AND A.ACNO = B.ACNO) WHERE A.NPAMAIN IN ('M','S','C','N','F')")
        oracle_execute_non_query("TEN", username, username, "COMMIT")

        ORACLE_CONN.Close()

        Dim SOL_SET As String
        Dim START_AMNT As Double
        Dim END_AMNT As Double
        Dim FILE_NAME As String


        Call GEN_NPA_XL_MACRO("ROKSD", 500000, 99999999999.99, "ROKSD")
        Call GEN_NPA_XL_MACRO("ROKSD", 300000, 499999.99, "ROKSD")
        Call GEN_NPA_XL_MACRO("ROKSD", 100000, 299999.99, "ROKSD")
        Call GEN_NPA_XL_MACRO("ROKSD", -99999999999.99, 99999.99, "ROKSD")


        Call GEN_NPA_XL_MACRO("ROKNR", 500000, 99999999999.99, "ROKNR")
        Call GEN_NPA_XL_MACRO("ROKNR", 300000, 499999.99, "ROKNR")
        Call GEN_NPA_XL_MACRO("ROKNR", 100000, 299999.99, "ROKNR")
        Call GEN_NPA_XL_MACRO("ROKNR", -99999999999.99, 99999.99, "ROKNR")


        Call GEN_NPA_XL_MACRO("ROTLY", 500000, 99999999999.99, "ROTLY")
        Call GEN_NPA_XL_MACRO("ROTLY", 300000, 499999.99, "ROTLY")
        Call GEN_NPA_XL_MACRO("ROTLY", 100000, 299999.99, "ROTLY")
        Call GEN_NPA_XL_MACRO("ROTLY", -99999999999.99, 99999.99, "ROTLY")


        Call GEN_NPA_XL_MACRO("ROKPT", 500000, 99999999999.99, "ROKPT")
        Call GEN_NPA_XL_MACRO("ROKPT", 300000, 499999.99, "ROKPT")
        Call GEN_NPA_XL_MACRO("ROKPT", 100000, 299999.99, "ROKPT")
        Call GEN_NPA_XL_MACRO("ROKPT", -99999999999.99, 99999.99, "ROKPT")


        Call GEN_NPA_XL_MACRO("ROTSR", 500000, 99999999999.99, "ROTSR")
        Call GEN_NPA_XL_MACRO("ROTSR", 300000, 499999.99, "ROTSR")
        Call GEN_NPA_XL_MACRO("ROTSR", 100000, 299999.99, "ROTSR")
        Call GEN_NPA_XL_MACRO("ROTSR", -99999999999.99, 99999.99, "ROTSR")


        Call GEN_NPA_XL_MACRO("ROKKD", 500000, 99999999999.99, "ROKKD")
        Call GEN_NPA_XL_MACRO("ROKKD", 300000, 499999.99, "ROKKD")
        Call GEN_NPA_XL_MACRO("ROKKD", 100000, 299999.99, "ROKKD")
        Call GEN_NPA_XL_MACRO("ROKKD", -99999999999.99, 99999.99, "ROKKD")


        Call GEN_NPA_XL_MACRO("ROMPM", 500000, 99999999999.99, "ROMPM")
        Call GEN_NPA_XL_MACRO("ROMPM", 300000, 499999.99, "ROMPM")
        Call GEN_NPA_XL_MACRO("ROMPM", 100000, 299999.99, "ROMPM")
        Call GEN_NPA_XL_MACRO("ROMPM", -99999999999.99, 99999.99, "ROMPM")


        Call GEN_NPA_XL_MACRO("ROEKM", 500000, 99999999999.99, "ROEKM")
        Call GEN_NPA_XL_MACRO("ROEKM", 300000, 499999.99, "ROEKM")
        Call GEN_NPA_XL_MACRO("ROEKM", 100000, 299999.99, "ROEKM")
        Call GEN_NPA_XL_MACRO("ROEKM", -99999999999.99, 99999.99, "ROEKM")


        Call GEN_NPA_XL_MACRO("ROKTM", 500000, 99999999999.99, "ROKTM")
        Call GEN_NPA_XL_MACRO("ROKTM", 300000, 499999.99, "ROKTM")
        Call GEN_NPA_XL_MACRO("ROKTM", 100000, 299999.99, "ROKTM")
        Call GEN_NPA_XL_MACRO("ROKTM", -99999999999.99, 99999.99, "ROKTM")


        Call GEN_NPA_XL_MACRO("ROTVM", 500000, 99999999999.99, "ROTVM")
        Call GEN_NPA_XL_MACRO("ROTVM", 300000, 499999.99, "ROTVM")
        Call GEN_NPA_XL_MACRO("ROTVM", 100000, 299999.99, "ROTVM")
        Call GEN_NPA_XL_MACRO("ROTVM", -99999999999.99, 99999.99, "ROTVM")


        Call GEN_NPA_XL_MACRO("ALL", 300000, 99999999999.99, "RL")   'R&L
        Call GEN_NPA_XL_MACRO("ALL", 100000, 299999.99, "CRM")    'CRM
        Call GEN_NPA_XL_MACRO("ALL", 300000, 99999999999.99, "GM")  'GM
        Call GEN_NPA_XL_MACRO("ALL", 500000, 99999999999.99, "CM")  'CHAIRMAN

        'Call GEN_NPA_XL("ALL", -9999999, 9999999999, "ALL")
        MessageBox.Show("FILE GENERATION OVER.")

        MsgBox("Excel file generation completed.", MsgBoxStyle.Information, "Process Completed")

        '=====================================================
        ''To ROs
        'NPA_THREAT_EXCEL_EMAIL("ROTVM", 500000, 99999999999.99, "aotvpm@yahoo.com", "ROTVM@keralagbank.com;kgb399@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTVM_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTVM_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTVM_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTVM_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROTVM)")
        'NPA_THREAT_EXCEL_EMAIL("ROEKM", 500000, 99999999999.99, "roekm.kgb@gmail.com", "ROEKM@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROEKM_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROEKM_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROEKM_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROEKM_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROEKM)")
        'NPA_THREAT_EXCEL_EMAIL("ROTSR", 500000, 99999999999.99, "smgbaotri@yahoo.co.in", "kgb400@keralagbank.com;ROKZD@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTSR_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTSR_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTSR_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTSR_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROTSR)")
        'NPA_THREAT_EXCEL_EMAIL("ROMPM", 500000, 99999999999.99, "kgbroperintalmanna@gmail.com", "ROPMA@keralagbank.com;kgb397@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROMPM_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROMPM_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROMPM_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROMPM_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROMPM)")
        'NPA_THREAT_EXCEL_EMAIL("ROKKD", 500000, 99999999999.99, "kgbroclt@gmail.com", "ROKZD@keralagbank.com;kgb393@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKKD_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKKD_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKKD_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKKD_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKKD)")
        'NPA_THREAT_EXCEL_EMAIL("ROTLY", 500000, 99999999999.99, "kgbrotly@gmail.com", "ROTLY@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTLY_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTLY_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTLY_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROTLY_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROTLY)")
        'NPA_THREAT_EXCEL_EMAIL("ROKSD", 500000, 99999999999.99, "kgbroksd@gmail.com", "ROKSD@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKSD_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKSD_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKSD_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKSD_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKSD)")
        'NPA_THREAT_EXCEL_EMAIL("ROKPT", 500000, 99999999999.99, "rokpta@gmail.com", "ROKPT@keralagbank.com;kgb393@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKPT_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKPT_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKPT_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKPT_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKPT)")
        'NPA_THREAT_EXCEL_EMAIL("ROKTM", 500000, 99999999999.99, "roktm.kgb@gmail.com", "kgb660@keralagbank.com;ROKTM@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKTM_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKTM_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKTM_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKTM_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKTM)")
        'NPA_THREAT_EXCEL_EMAIL("ROKNR", 500000, 99999999999.99, "nmgbknrao@gmail.com", "ROKNR@keralagbank.com;kgb394@keralagbank.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKNR_500000_99999999999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKNR_300000_499999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKNR_100000_299999.99.xlsx", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_ROKNR_Upto_99999.99.xlsx", "NPA Threat for the next 7 Days (SOLSET : ROKNR)")

        ''To GMs
        'NPA_THREAT_EXCEL_EMAIL("ALL", 300000, 99999999999.99, "thangavelu.ppt@gmail.com;nkkrishnankutty46876@gmail.com;haridasanv@gmail.com;srnair32474@gmail.com", "", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_GM_300000_99999999999.99.xlsx", "0", "0", "0", "NPA Threat for the next 7 Days (SOLSET : ALL; Outstanding from 300000.00 to 9999999999.99)")

        ''To CRM
        'NPA_THREAT_EXCEL_EMAIL("ALL", 100000, 299999.99, "crmwing.kgb@gmail.com", "", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_CRM_100000_299999.99.xlsx", "0", "0", "0", "NPA Threat for the next 7 Days (SOLSET : ALL; Outstanding from 100000.00 to 299999.99)")

        ''To RL
        'NPA_THREAT_EXCEL_EMAIL("ALL", 300000, 99999999999.99, "smgbrl93@gmail.com", "", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_RL_300000_99999999999.99.xlsx", "0", "0", "0", "NPA Threat for the next 7 Days (SOLSET : ALL; Outstanding from 300000.00 to 9999999999.99)")

        ''To CHM
        'NPA_THREAT_EXCEL_EMAIL("ALL", 500000, 99999999999.99, "chairmankeralagb@gmail.com", "franklinkf@gmail.com;sureshsmgb1@gmail.com;udayakumarcv@gmail.com", "D:\PNPA\7DAY_NPA_THREAT_" & Format(Today, "ddMMyyyy") & "_CM_500000_99999999999.99.xlsx", "0", "0", "0", "NPA Threat for the next 7 Days (SOLSET : ALL; Outstanding from 500000.00 to 9999999999.99)")

        'MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")
        '=====================================================

    End Sub

    Sub sendemail_pnpa(ByVal sendfromaccount As String, ByVal database As String, ByVal user As String, ByVal password As String)

        ''Generating EMail
        ''Add "Microsoft Outlook 15.0 Object Library" in Project >> Reference >> Com
        ''Add the following in declaration part
        ''Imports System.Runtime.InteropServices
        ''Imports Outlook = Microsoft.Office.Interop.Outlook

        Dim oradb As String = "Data Source=" & database & ";User Id= " & user & ";Password= " & password & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Sending Mail")

        Dim sql1 As String
        Dim sql2 As String
        Dim tmplength As Integer = 0
        Dim substringno As Integer = 0
        Dim substringstart As Integer = 1
        Dim Bodydata As String = ""

        tempcount = 0
        sql = "SELECT MAIL_DATAID,MAIL_DATASUBID,MAIL_TO,NVL(MAIL_CC,'A')MAIL_CC,NVL(MAIL_BCC,'A')MAIL_BCC,MAIL_SUBJECT FROM Z_EMAIL ORDER BY MAIL_DATAID, MAIL_DATASUBID"
        Dim cmd21 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd21.ExecuteReader()
        While dr.Read()
            Bodydata = ""
            tmplength = 0
            substringno = 0
            substringstart = 1

            sql1 = "SELECT LENGTH(MAIL_BODY) TEMPLENGTH FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
            Dim cmd22 As New OracleCommand(sql1, conn)
            Dim dr1 As OracleDataReader = cmd22.ExecuteReader()
            While dr1.Read()
                tmplength = dr1.Item("TEMPLENGTH")
                Do Until tmplength = 0

                    If tmplength >= 3000 Then

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & ",3000) ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()

                    Else

                        substringno = tmplength

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & "," & substringno & ") ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()
                        tmplength = 0

                    End If

                Loop

            End While
            dr1.Close()

            Dim oApp As Outlook._Application
            oApp = New Outlook.Application()
            Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
            Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, sendfromaccount)
            newMail.To = dr.Item("MAIL_TO")
            If dr.Item("MAIL_CC") <> "A" Then
                newMail.CC = dr.Item("MAIL_CC")
            End If
            If dr.Item("MAIL_BCC") <> "A" Then
                newMail.BCC = dr.Item("MAIL_BCC")
            End If

            If dr.Item("MAIL_DATASUBID") = "ALL1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_CM_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "CRM" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_CRM_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "RL" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_RL_300000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "GMTHAN1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_GM_300000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "GMKRISH1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_GM_300000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "GMHARI1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_GM_300000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "GMRADH1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_GM_300000_99999999999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROTVM1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTVM_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTVM2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTVM_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTVM3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTVM_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTVM4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTVM_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROEKM1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROEKM_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROEKM2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROEKM_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROEKM3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROEKM_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROEKM4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROEKM_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROTSR1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTSR_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTSR2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTSR_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTSR3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTSR_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTSR4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTSR_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROMPM1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROMPM_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROMPM2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROMPM_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROMPM3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROMPM_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROMPM4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROMPM_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROKKD1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKKD_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKKD2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKKD_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKKD3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKKD_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKKD4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKKD_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROTLY1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTLY_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTLY2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTLY_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTLY3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTLY_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROTLY4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROTLY_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROKSD1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKSD_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKSD2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKSD_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKSD3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKSD_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKSD4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKSD_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROKPT1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKPT_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKPT2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKPT_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKPT3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKPT_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKPT4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKPT_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROKTM1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKTM_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKTM2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKTM_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKTM3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKTM_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKTM4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKTM_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROKTM1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKTM_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKTM2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKTM_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKTM3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKTM_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKTM4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKTM_Upto_99999.99.xlsx")
            End If

            If dr.Item("MAIL_DATASUBID") = "ROKNR1" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKNR_500000_99999999999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKNR2" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKNR_300000_499999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKNR3" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKNR_100000_299999.99.xlsx")
            End If
            If dr.Item("MAIL_DATASUBID") = "ROKNR4" Then
                newMail.Attachments.Add("D:\PNPA\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_ROKNR_Upto_99999.99.xlsx")
            End If

            newMail.Subject = dr.Item("MAIL_SUBJECT")
            newMail.HTMLBody = Bodydata
            newMail.SendUsingAccount = account
            newMail.Send()
            tempcount = tempcount + 1
            processmessage("Sending Mail No - " & tempcount)

        End While
        dr.Close()

        conn.Close()
        conn.Dispose()

        processmessage("")

        MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")

    End Sub

    Private Sub NPA_THREAT_EXCEL_EMAIL(ByVal SOLSET As String, ByVal FROM_AMT As Double, ByVal TO_AMT As Double, ByVal SEND_TO As String, ByVal SEND_CC As String, ByVal ATTACH1 As String, ByVal ATTACH2 As String, ByVal ATTACH3 As String, ByVal ATTACH4 As String, ByVal SUBJECT As String)

        'Generating EMail
        'Add "Microsoft Outlook 15.0 Object Library" in Project >> Reference >> Com
        'Add the following in declaration part
        'Imports System.Runtime.InteropServices
        'Imports Outlook = Microsoft.Office.Interop.Outlook

        Dim oApp As Outlook._Application
        oApp = New Outlook.Application()
        Dim outlooksendfromaccount As String
        Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        Dim dirs As String() = Directory.GetFiles("D:\PNPA")
        Dim dir As String

        outlooksendfromaccount = "smgbmis2@gmail.com"

        Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)

        newMail.To = SEND_TO
        newMail.CC = SEND_CC
        newMail.Subject = SUBJECT
        newMail.HTMLBody = "<html><body><head><style>.normalandleft{TEXT-ALIGN: left; FONT-FAMILY: arial, helvetica, sans-serif; FONT-SIZE: 9pt; FONT-WEIGHT: normal;}</style></head><p class=normalandleft>Dear Sir,</p><p class=normalandleft>Enclosed please find the attachment containing the subject data in excel file/s.  Master data is displayed in two sheets, one sorted based on indicative NPA date and the other sorted based on SOLID. Data in the excel file is setup and aligned to print directly in A4 size paper in Landscape mode.  At the end of each account, one line is added to note your remarks after making the follow up.</p><p class=normalandleft>Please write to misteam.kgb@gmail.com for any issues/suggestions.</p><p class=normalandleft>Yours faithfully,</p><p class=normalandleft>MIS Team</p></body></html>"
        If ATTACH1 <> "0" Then
            newMail.Attachments.Add(ATTACH1)
        End If
        If ATTACH2 <> "0" Then
            newMail.Attachments.Add(ATTACH2)
        End If
        If ATTACH3 <> "0" Then
            newMail.Attachments.Add(ATTACH3)
        End If
        If ATTACH4 <> "0" Then
            newMail.Attachments.Add(ATTACH4)
        End If
        newMail.SendUsingAccount = account
        newMail.Send()

    End Sub

    Private Sub GEN_NPA_XL(ByVal SOLSET As String, ByVal START_AMT As Double, ByVal END_AMT As Double, ByVal FILENAME As String)  'NPA FOLLOWUP XL DATA GENERATION
        Dim WORK_BOOK_NAME As String
        If START_AMT > 0 Then
            WORK_BOOK_NAME = "7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_" & FILENAME & "_" & Trim(START_AMT) & "_" & Trim(END_AMT) & ".xlsx"
        Else
            WORK_BOOK_NAME = "7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_" & FILENAME & "_" & "Upto" & "_" & Trim(END_AMT) & ".xlsx"
        End If
        SOLSET = "'" & SOLSET & "'"

        '==============================================

        Dim oracle_cnn_string As String = "Data Source=ten; User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        '================= CREATING EXCEL FILE ================================

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return
        End If
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim formatRange As Excel.Range
        Dim cell As Excel.Range
        Dim border As Excel.Borders

        xlWorkBook = xlApp.Workbooks.Add(misValue)

        Dim sql1 As String

        sql1 = "SELECT ACNO,SCHEMECODE,SOLID,NVL(DATE1,'01-JAN-1901'),NVL(DATE2,'01-JAN-1901'),NVL(DATE4,'01-JAN-1901'),NVL(DATE11,'01-JAN-1901'),NVL(DATE12,'01-JAN-1901'),NVL(DATE13,'01-JAN-1901'),NVL(NUMBER1,0),NVL(NUMBER2,0),NVL(NUMBER3,0),NVL(NUMBER4,0),NVL(NUMBER18,0),NVL(TEXT1,' '),NVL(TEXT2,' '),NVL(TEXT3,' '),NVL(TEXT4,' '),NVL(TEXT5,' '),NVL(TEXT7,' '),NVL(TEXT11,' '),NVL(TEXT12,' '),NVL(TEXT13,' '),NVL(TEXT14,' '), NUMBER20, DATE20, NUMBER21, NVL(TEXT6,' '), NVL(NUMBER5,0) FROM C_MISADV WHERE ( NUMBER20>0 ) AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO, NUMBER20, TEXT1"
        Dim cmd1 As New OracleCommand(sql1, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()


        '================= BEGINNING OF SHEET 1 ==========================================================================================================

        '================= NAMING SHEET 1 ================================
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        xlWorkSheet.Name = "All In One - Date Wise"

        Dim PREVDT As String = ""
        Dim PREVACNO As String = ""


        '================= SETTING COLUMN WIDTHS ================================

        xlWorkSheet.Range("A1").ColumnWidth = 9
        xlWorkSheet.Range("B1").ColumnWidth = 13
        xlWorkSheet.Range("C1").ColumnWidth = 11
        xlWorkSheet.Range("D1").ColumnWidth = 18
        xlWorkSheet.Range("E1").ColumnWidth = 9
        xlWorkSheet.Range("F1").ColumnWidth = 9
        xlWorkSheet.Range("G1").ColumnWidth = 13
        xlWorkSheet.Range("H1").ColumnWidth = 11
        xlWorkSheet.Range("I1").ColumnWidth = 10
        xlWorkSheet.Range("J1").ColumnWidth = 22
        xlWorkSheet.Range("K1").ColumnWidth = 10
        xlWorkSheet.Range("L1").ColumnWidth = 10

        '================= Setting Column Headings ================================
        xlWorkSheet.Cells(1, 1) = "NPA THREAT FOR THE NEXT 7 DAYS AS ON " & Format(RptDate, "dd-MM-yyyy")
        xlWorkSheet.Cells(2, 1) = "SOLSET : " & Replace(SOLSET, "'", "") & " (OUTSTANDING BETWEEN " & Trim(START_AMT) & " AND " & Trim(END_AMT) & ")"

        xlWorkSheet.Range("A1:L1").Merge()
        xlWorkSheet.Range("A1:L1").HorizontalAlignment = 3
        xlWorkSheet.Range("A1:L1").Font.Name = "Arial"
        xlWorkSheet.Range("A1:L1").Font.Size = 11
        xlWorkSheet.Range("A1:L1").Font.Bold = True


        xlWorkSheet.Range("A2:L2").Merge()
        xlWorkSheet.Range("A2:L2").HorizontalAlignment = 3
        xlWorkSheet.Range("A2:L2").Font.Name = "Arial"
        xlWorkSheet.Range("A2:L2").Font.Size = 10
        xlWorkSheet.Range("A2:L2").Font.Bold = True


        xlWorkSheet.Cells(3, 3) = "Branch"
        xlWorkSheet.Cells(3, 4) = "Branch Name"
        xlWorkSheet.Cells(3, 5) = "Open Date"
        xlWorkSheet.Cells(3, 6) = "Online Date"
        xlWorkSheet.Cells(3, 7) = "RO/DT"
        xlWorkSheet.Cells(3, 8) = "CUG - Office"
        xlWorkSheet.Cells(3, 9) = "CUG - Mobile"
        xlWorkSheet.Cells(3, 10) = "Manager"
        xlWorkSheet.Cells(3, 11) = "Since"
        xlWorkSheet.Cells(3, 12) = "Staff Strength"

        xlWorkSheet.Cells(4, 3) = "Scheme"
        xlWorkSheet.Cells(4, 4) = "Open Date"
        xlWorkSheet.Cells(4, 5) = "Loan"
        xlWorkSheet.Cells(4, 6) = "Bal O/S"
        xlWorkSheet.Cells(4, 7) = "Due Date"
        xlWorkSheet.Cells(4, 8) = "Overdue"
        xlWorkSheet.Cells(4, 9) = "Critical Amt"
        xlWorkSheet.Cells(4, 10) = "Cri.Amt.Qtr.End"
        xlWorkSheet.Cells(4, 11) = "NPA Reason"

        xlWorkSheet.Cells(4, 12) = "AOD Due On"


        xlWorkSheet.Cells(5, 3) = "Rep Schedule"
        xlWorkSheet.Cells(5, 4) = "Rep Period"
        xlWorkSheet.Cells(5, 5) = "Installment"
        xlWorkSheet.Cells(5, 6) = "Rep Frequency"
        xlWorkSheet.Cells(5, 8) = "First Inst Date"
        xlWorkSheet.Cells(5, 9) = "Whether Rescheduled In Finacle"

        xlWorkSheet.Cells(6, 3) = "Parties"
        xlWorkSheet.Cells(6, 4) = "Cust ID"
        xlWorkSheet.Cells(6, 5) = "Customer Name"
        xlWorkSheet.Cells(6, 8) = "Relation"
        xlWorkSheet.Cells(6, 10) = "Gold Loan"
        xlWorkSheet.Cells(6, 11) = "Mobile No"
        xlWorkSheet.Cells(6, 12) = "TotalDep"

        xlWorkSheet.Cells(7, 3) = "Notice"
        xlWorkSheet.Cells(7, 4) = "Notice Date"
        xlWorkSheet.Cells(7, 5) = "Send To"
        xlWorkSheet.Cells(7, 8) = "Notice Name"

        xlWorkSheet.Cells(8, 1) = "NPA dt."
        xlWorkSheet.Cells(8, 2) = "A/c No"
        xlWorkSheet.Cells(8, 3) = "Follow Up"
        xlWorkSheet.Cells(8, 4) = "Date"
        xlWorkSheet.Cells(8, 5) = "Contacted"
        xlWorkSheet.Cells(8, 6) = "Cont Type"
        xlWorkSheet.Cells(8, 7) = "Initiated By"
        xlWorkSheet.Cells(8, 8) = "Done By"
        xlWorkSheet.Cells(8, 9) = "Response"

        formatRange = xlWorkSheet.Range("A3", "L8")
        formatRange.HorizontalAlignment = 3
        formatRange.Font.Bold = True
        formatRange.Font.Size = 9



        xlWorkSheet.Range("F5", "G5").Merge()
        xlWorkSheet.Range("I5", "L5").Merge()
        xlWorkSheet.Range("E6", "G6").Merge()
        xlWorkSheet.Range("H6", "I6").Merge()
        xlWorkSheet.Range("E7", "G7").Merge()
        xlWorkSheet.Range("H7", "L7").Merge()
        xlWorkSheet.Range("I8", "L8").Merge()

        Dim CRITROW As Long

        Dim npamainnum As Integer

        Dim acno As String
        Dim Contacted As String
        Dim ContType As String
        Dim cugland As String
        Dim cugmob As String
        Dim CustID As String
        Dim CustomerName As String
        Dim DATE1 As Date
        Dim DATE11 As Date
        Dim DATE12 As Date
        Dim DATE13 As Date
        Dim DATE2 As Date
        Dim DATE4 As Date
        Dim DoneBy As String
        Dim FirstInstDate As Date
        Dim FUDate As Date
        Dim GoldLoan As String
        Dim InitiatedBy As String
        Dim Installment As String
        Dim MobileNo As String
        Dim NoticeDate As Date
        Dim NoticeName As String
        Dim NUMBER1 As Double
        Dim NUMBER2 As Double
        Dim NUMBER3 As Double
        Dim NUMBER4 As Double
        Dim Relation As String
        Dim RePeriod As Double
        Dim RepFrequency As String
        Dim Response As String
        Dim RO_DT As String
        Dim SCHEMECODE As String
        Dim SendTo As String
        Dim solid As String
        Dim STFPOS As Integer
        Dim TEXT1 As String
        Dim TEXT11 As String
        Dim TEXT14 As String
        Dim TotalDep As Double
        Dim WhetherRescheduled As String
        Dim NPADT As Date
        Dim CRIAMTQTREND As Double

        Dim ROWNUM As Integer
        ROWNUM = 8

        '================= WRITING DATA TO EXCEL ================================
        Dim totnum As Integer = 0
        Dim totbal As Double = 0
        Dim totcritic As Double = 0

        While dr1.Read

            npamainnum = dr1(24)

            Select Case npamainnum

                Case 1
                    acno = dr1(0)
                    solid = dr1(2)
                    TEXT11 = dr1(20)
                    DATE11 = dr1(6)
                    DATE12 = dr1(7)
                    RO_DT = dr1(21) & "/" & dr1(22)
                    cugland = Trim(Val(solid) - 40000 + 4000)
                    cugmob = Trim(Val(solid) - 40000 + 9400999000)
                    TEXT14 = dr1(23)
                    DATE13 = dr1(8)
                    STFPOS = dr1(13)
                    NPADT = dr1(25)

                    ROWNUM = ROWNUM + 1

                    If ROWNUM <> 9 Then
                        xlWorkSheet.Cells(ROWNUM, 1) = PREVDT
                        xlWorkSheet.Cells(ROWNUM, 2) = PREVACNO
                        xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                        xlWorkSheet.Cells(ROWNUM, 3) = "Remarks"
                        xlWorkSheet.Range("D" & ROWNUM, "L" & ROWNUM).Merge()
                        ROWNUM = ROWNUM + 1
                    End If

                    processmessage(WORK_BOOK_NAME & ":All in one : Writing row " & ROWNUM)
                    Application.DoEvents()

                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = solid
                    xlWorkSheet.Cells(ROWNUM, 4) = UCase(TEXT11)
                    xlWorkSheet.Cells(ROWNUM, 5) = DATE11
                    xlWorkSheet.Cells(ROWNUM, 6) = DATE12
                    xlWorkSheet.Cells(ROWNUM, 7) = RO_DT
                    xlWorkSheet.Cells(ROWNUM, 8) = cugland
                    xlWorkSheet.Cells(ROWNUM, 9) = cugmob
                    xlWorkSheet.Cells(ROWNUM, 10) = TEXT14
                    xlWorkSheet.Cells(ROWNUM, 11) = DATE13
                    xlWorkSheet.Cells(ROWNUM, 12) = STFPOS
                    PREVDT = NPADT
                    PREVACNO = acno


                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("F" & ROWNUM, "F" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("G" & ROWNUM, "G" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("I" & ROWNUM, "I" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("J" & ROWNUM, "J" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("K" & ROWNUM, "K" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("L" & ROWNUM, "L" & ROWNUM).NumberFormat = "@"

                    xlWorkSheet.Range("A" & ROWNUM, "L" & ROWNUM).Font.Bold = True

                    acno = dr1(0)
                    SCHEMECODE = dr1(15)
                    DATE1 = dr1(3)
                    NUMBER1 = dr1(9)
                    NUMBER2 = dr1(10)
                    DATE2 = dr1(4)
                    NUMBER3 = dr1(11)
                    NUMBER4 = dr1(12)
                    TEXT1 = dr1(14)
                    DATE4 = dr1(5)
                    CRIAMTQTREND = Math.Round(dr1(28), 0, MidpointRounding.AwayFromZero)


                    ROWNUM = ROWNUM + 1

                    totnum = totnum + 1
                    totbal = totbal + NUMBER2
                    totcritic = totcritic + NUMBER4



                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = SCHEMECODE
                    xlWorkSheet.Cells(ROWNUM, 4) = DATE1
                    xlWorkSheet.Cells(ROWNUM, 5) = NUMBER1
                    xlWorkSheet.Cells(ROWNUM, 6) = NUMBER2
                    xlWorkSheet.Cells(ROWNUM, 7) = DATE2
                    xlWorkSheet.Cells(ROWNUM, 8) = NUMBER3
                    xlWorkSheet.Cells(ROWNUM, 9) = NUMBER4

                    CRITROW = ROWNUM
                    xlWorkSheet.Cells(ROWNUM, 10) = CRIAMTQTREND
                    xlWorkSheet.Cells(ROWNUM, 11) = TEXT1
                    xlWorkSheet.Cells(ROWNUM, 12) = DATE4

                    If NUMBER4 > 0 Then
                        xlWorkSheet.Range("I" & CRITROW).AddComment()
                        xlWorkSheet.Range("I" & CRITROW).Comment.Text("Balance to Crit.amt ratio=" & Format(NUMBER2 / NUMBER4, "fixed"))
                    End If

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("F" & ROWNUM, "F" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("G" & ROWNUM, "G" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("I" & ROWNUM, "I" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("J" & ROWNUM, "J" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("J" & ROWNUM, "K" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("K" & ROWNUM, "L" & ROWNUM).NumberFormat = "DD-MM-YYYY"

                Case 2  'S
                    acno = dr1(0)
                    RePeriod = dr1(9)
                    Installment = dr1(10)
                    RepFrequency = dr1(14)
                    FirstInstDate = dr1(3)
                    WhetherRescheduled = dr1(15)

                    ROWNUM = ROWNUM + 1

                    xlWorkSheet.Range("F" & ROWNUM, "G" & ROWNUM).Merge()
                    xlWorkSheet.Range("I" & ROWNUM, "L" & ROWNUM).Merge()

                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = "Rep Schedule"
                    xlWorkSheet.Cells(ROWNUM, 4) = RePeriod
                    xlWorkSheet.Cells(ROWNUM, 5) = Installment
                    xlWorkSheet.Cells(ROWNUM, 6) = RepFrequency
                    xlWorkSheet.Cells(ROWNUM, 8) = FirstInstDate
                    xlWorkSheet.Cells(ROWNUM, 9) = WhetherRescheduled

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("F" & ROWNUM, "F" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("I" & ROWNUM, "I" & ROWNUM).NumberFormat = "@"

                Case 3   'C
                    acno = dr1(0)
                    CustID = dr1(18)
                    CustomerName = dr1(15)
                    Relation = dr1(14)
                    GoldLoan = dr1(10)
                    MobileNo = dr1(16)
                    TotalDep = dr1(9)

                    ROWNUM = ROWNUM + 1

                    xlWorkSheet.Range("E" & ROWNUM, "G" & ROWNUM).Merge()
                    xlWorkSheet.Range("H" & ROWNUM, "I" & ROWNUM).Merge()

                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = "Parties"
                    xlWorkSheet.Cells(ROWNUM, 4) = CustID
                    xlWorkSheet.Cells(ROWNUM, 5) = CustomerName
                    xlWorkSheet.Cells(ROWNUM, 8) = Relation
                    xlWorkSheet.Cells(ROWNUM, 10) = GoldLoan
                    xlWorkSheet.Cells(ROWNUM, 11) = MobileNo
                    xlWorkSheet.Cells(ROWNUM, 12) = TotalDep

                    If Relation = "A/C HOLDER" Then
                        If NUMBER4 > 0 Then
                            If TotalDep >= NUMBER4 Then
                                xlWorkSheet.Range("I" & CRITROW).Interior.Color = 13434825
                                xlWorkSheet.Range("I" & CRITROW).Font.Bold = True
                            End If
                        End If

                        If CRIAMTQTREND > 0 Then
                            If TotalDep >= CRIAMTQTREND Then
                                xlWorkSheet.Range("J" & CRITROW).Interior.Color = 13434825
                                xlWorkSheet.Range("J" & CRITROW).Font.Bold = True
                            End If
                        End If
                    End If

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("J" & ROWNUM, "J" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("K" & ROWNUM, "K" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("L" & ROWNUM, "L" & ROWNUM).NumberFormat = "##########"


                Case 4   'N
                    acno = dr1(0)
                    NoticeDate = dr1(3)
                    SendTo = dr1(14)
                    NoticeName = dr1(15)

                    ROWNUM = ROWNUM + 1

                    xlWorkSheet.Range("E" & ROWNUM, "G" & ROWNUM).Merge()
                    xlWorkSheet.Range("H" & ROWNUM, "L" & ROWNUM).Merge()

                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = "Notice"
                    xlWorkSheet.Cells(ROWNUM, 4) = NoticeDate
                    xlWorkSheet.Cells(ROWNUM, 5) = SendTo
                    xlWorkSheet.Cells(ROWNUM, 8) = NoticeName

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("K" & ROWNUM, "K" & ROWNUM).NumberFormat = "@"



                Case 5   'F
                    acno = dr1(0)
                    FUDate = dr1(3)
                    Contacted = dr1(18)

                    If Contacted = "PARTY" Then Contacted = Contacted & "-" & dr1(27)

                    ContType = dr1(15)
                    InitiatedBy = dr1(14)
                    DoneBy = dr1(17)
                    Response = dr1(19)


                    ROWNUM = ROWNUM + 1

                    xlWorkSheet.Range("I" & ROWNUM, "L" & ROWNUM).Merge()


                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = "FollowUp"
                    xlWorkSheet.Cells(ROWNUM, 4) = FUDate
                    xlWorkSheet.Cells(ROWNUM, 5) = Contacted
                    xlWorkSheet.Cells(ROWNUM, 6) = ContType
                    xlWorkSheet.Cells(ROWNUM, 7) = InitiatedBy
                    xlWorkSheet.Cells(ROWNUM, 8) = DoneBy
                    xlWorkSheet.Cells(ROWNUM, 9) = Response

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("J" & ROWNUM, "J" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("K" & ROWNUM, "K" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("L" & ROWNUM, "L" & ROWNUM).NumberFormat = "@"


                Case Else

            End Select
        End While
        dr1.Close()

        '=============================



        '=================================


        xlWorkSheet.Range("A9").Select()
        xlWorkSheet.Application.ActiveWindow.FreezePanes = True


        '================= SETTING FONT SIZE OF DATA ================================
        formatRange = xlWorkSheet.Range("A3", "L" & ROWNUM)
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"


        xlWorkSheet.Range("A3:B7").Merge()

        xlWorkSheet.Range("A3").HorizontalAlignment = 2
        xlWorkSheet.Range("A3").VerticalAlignment = 2
        xlWorkSheet.Range("A3").FormulaR1C1 = "No of Accounts : " & totnum & Chr(10) & "Bal O/s : " & Format(totbal / 100000, "fixed") & " Lakhs" & Chr(10) & "Crit Amt : " & Format(totcritic / 100000, "fixed") & " Lakhs"
        xlWorkSheet.Range("A3").Font.Bold = True
        xlWorkSheet.Range("A3").Font.Size = 11


        formatRange = xlWorkSheet.Range("A9", "L" & ROWNUM)
        formatRange.HorizontalAlignment = 2


        '==============PAGE & PRINT SETUP=================================

        With xlWorkSheet.PageSetup
            .PrintTitleRows = "$3:$8"
            .PrintTitleColumns = ""
        End With

        With xlWorkSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""

            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.05)
            .FooterMargin = xlApp.InchesToPoints(0.05)

            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With

        '================= END OF SHEET 1 ==========================================================================================================


        Dim SQLX As String = "SELECT ACNO,SCHEMECODE,SOLID,NVL(DATE1,'01-JAN-1901'),NVL(DATE2,'01-JAN-1901'),NVL(DATE4,'01-JAN-1901'),NVL(DATE11,'01-JAN-1901'),NVL(DATE12,'01-JAN-1901'),NVL(DATE13,'01-JAN-1901'),NVL(NUMBER1,0),NVL(NUMBER2,0),NVL(NUMBER3,0),NVL(NUMBER4,0),NVL(NUMBER18,0),NVL(TEXT1,' '),NVL(TEXT2,' '),NVL(TEXT3,' '),NVL(TEXT4,' '),NVL(TEXT5,' '),NVL(TEXT7,' '),NVL(TEXT11,' '),NVL(TEXT12,' '),NVL(TEXT13,' '),NVL(TEXT14,' '), NUMBER20, DATE20, NUMBER21, NVL(TEXT6,' '), NVL(NUMBER5,0) FROM C_MISADV WHERE ( NUMBER20>0 ) AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY SOLID, DATE20, NUMBER21 DESC, ACNO, NUMBER20, TEXT1"
        Dim cmdx As New OracleCommand(SQLX, oracle_conn)
        Dim drx As OracleDataReader = cmdx.ExecuteReader()

        '================= BEGINNING OF SHEET 2 ==========================================================================================================

        '================= NAMING SHEET 2 ================================
        Dim xlNewSheet = xlWorkBook.Worksheets.Add(After:=xlWorkBook.Worksheets(1))
        xlWorkSheet = xlNewSheet
        xlWorkSheet.Name = "All In One - SOL Wise"




        '================= SETTING COLUMN WIDTHS ================================

        xlWorkSheet.Range("A1").ColumnWidth = 9
        xlWorkSheet.Range("B1").ColumnWidth = 13
        xlWorkSheet.Range("C1").ColumnWidth = 11
        xlWorkSheet.Range("D1").ColumnWidth = 18
        xlWorkSheet.Range("E1").ColumnWidth = 9
        xlWorkSheet.Range("F1").ColumnWidth = 9
        xlWorkSheet.Range("G1").ColumnWidth = 13
        xlWorkSheet.Range("H1").ColumnWidth = 11
        xlWorkSheet.Range("I1").ColumnWidth = 10
        xlWorkSheet.Range("J1").ColumnWidth = 22
        xlWorkSheet.Range("K1").ColumnWidth = 10
        xlWorkSheet.Range("L1").ColumnWidth = 10

        '================= Setting Column Headings ================================
        xlWorkSheet.Cells(1, 1) = "NPA THREAT FOR THE NEXT 7 DAYS AS ON " & Format(RptDate, "dd-MM-yyyy")
        xlWorkSheet.Cells(2, 1) = "SOLSET : " & Replace(SOLSET, "'", "") & " (OUTSTANDING BETWEEN " & Trim(START_AMT) & " AND " & Trim(END_AMT) & ")"

        xlWorkSheet.Range("A1:L1").Merge()
        xlWorkSheet.Range("A1:L1").HorizontalAlignment = 3
        xlWorkSheet.Range("A1:L1").Font.Name = "Arial"
        xlWorkSheet.Range("A1:L1").Font.Size = 11
        xlWorkSheet.Range("A1:L1").Font.Bold = True


        xlWorkSheet.Range("A2:L2").Merge()
        xlWorkSheet.Range("A2:L2").HorizontalAlignment = 3
        xlWorkSheet.Range("A2:L2").Font.Name = "Arial"
        xlWorkSheet.Range("A2:L2").Font.Size = 10
        xlWorkSheet.Range("A2:L2").Font.Bold = True


        xlWorkSheet.Cells(3, 3) = "Branch"
        xlWorkSheet.Cells(3, 4) = "Branch Name"
        xlWorkSheet.Cells(3, 5) = "Open Date"
        xlWorkSheet.Cells(3, 6) = "Online Date"
        xlWorkSheet.Cells(3, 7) = "RO/DT"
        xlWorkSheet.Cells(3, 8) = "CUG - Office"
        xlWorkSheet.Cells(3, 9) = "CUG - Mobile"
        xlWorkSheet.Cells(3, 10) = "Manager"
        xlWorkSheet.Cells(3, 11) = "Since"
        xlWorkSheet.Cells(3, 12) = "Staff Strength"

        xlWorkSheet.Cells(4, 3) = "Scheme"
        xlWorkSheet.Cells(4, 4) = "Open Date"
        xlWorkSheet.Cells(4, 5) = "Loan"
        xlWorkSheet.Cells(4, 6) = "Bal O/S"
        xlWorkSheet.Cells(4, 7) = "Due Date"
        xlWorkSheet.Cells(4, 8) = "Overdue"
        xlWorkSheet.Cells(4, 9) = "Critical Amt"
        xlWorkSheet.Cells(4, 10) = "Cri.Amt.Qtr.End"
        xlWorkSheet.Cells(4, 11) = "NPA Reason"

        xlWorkSheet.Cells(4, 12) = "AOD Due On"


        xlWorkSheet.Cells(5, 3) = "Rep Schedule"
        xlWorkSheet.Cells(5, 4) = "Rep Period"
        xlWorkSheet.Cells(5, 5) = "Installment"
        xlWorkSheet.Cells(5, 6) = "Rep Frequency"
        xlWorkSheet.Cells(5, 8) = "First Inst Date"
        xlWorkSheet.Cells(5, 9) = "Whether Rescheduled In Finacle"

        xlWorkSheet.Cells(6, 3) = "Parties"
        xlWorkSheet.Cells(6, 4) = "Cust ID"
        xlWorkSheet.Cells(6, 5) = "Customer Name"
        xlWorkSheet.Cells(6, 8) = "Relation"
        xlWorkSheet.Cells(6, 10) = "Gold Loan"
        xlWorkSheet.Cells(6, 11) = "Mobile No"
        xlWorkSheet.Cells(6, 12) = "TotalDep"

        xlWorkSheet.Cells(7, 3) = "Notice"
        xlWorkSheet.Cells(7, 4) = "Notice Date"
        xlWorkSheet.Cells(7, 5) = "Send To"
        xlWorkSheet.Cells(7, 8) = "Notice Name"

        xlWorkSheet.Cells(8, 1) = "NPA dt."
        xlWorkSheet.Cells(8, 2) = "A/c No"
        xlWorkSheet.Cells(8, 3) = "Follow Up"
        xlWorkSheet.Cells(8, 4) = "Date"
        xlWorkSheet.Cells(8, 5) = "Contacted"
        xlWorkSheet.Cells(8, 6) = "Cont Type"
        xlWorkSheet.Cells(8, 7) = "Initiated By"
        xlWorkSheet.Cells(8, 8) = "Done By"
        xlWorkSheet.Cells(8, 9) = "Response"

        formatRange = xlWorkSheet.Range("A3", "L8")
        formatRange.HorizontalAlignment = 3
        formatRange.Font.Bold = True
        formatRange.Font.Size = 9



        xlWorkSheet.Range("F5", "G5").Merge()
        xlWorkSheet.Range("I5", "L5").Merge()
        xlWorkSheet.Range("E6", "G6").Merge()
        xlWorkSheet.Range("H6", "I6").Merge()
        xlWorkSheet.Range("E7", "G7").Merge()
        xlWorkSheet.Range("H7", "L7").Merge()
        xlWorkSheet.Range("I8", "L8").Merge()


        ROWNUM = 8

        '================= WRITING DATA TO EXCEL ================================
        totnum = 0
        totbal = 0
        totcritic = 0



        While drx.Read

            npamainnum = drx(24)

            Select Case npamainnum

                Case 1
                    acno = drx(0)
                    solid = drx(2)
                    TEXT11 = drx(20)
                    DATE11 = drx(6)
                    DATE12 = drx(7)
                    RO_DT = drx(21) & "/" & drx(22)
                    cugland = Trim(Val(solid) - 40000 + 4000)
                    cugmob = Trim(Val(solid) - 40000 + 9400999000)
                    TEXT14 = drx(23)
                    DATE13 = drx(8)
                    STFPOS = drx(13)
                    NPADT = drx(25)

                    ROWNUM = ROWNUM + 1

                    If ROWNUM <> 9 Then
                        xlWorkSheet.Cells(ROWNUM, 1) = PREVDT
                        xlWorkSheet.Cells(ROWNUM, 2) = PREVACNO
                        xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                        xlWorkSheet.Cells(ROWNUM, 3) = "Remarks"
                        xlWorkSheet.Range("D" & ROWNUM, "L" & ROWNUM).Merge()
                        ROWNUM = ROWNUM + 1
                    End If

                    processmessage(WORK_BOOK_NAME & ":All in one : Writing row " & ROWNUM)
                    Application.DoEvents()

                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = solid
                    xlWorkSheet.Cells(ROWNUM, 4) = UCase(TEXT11)
                    xlWorkSheet.Cells(ROWNUM, 5) = DATE11
                    xlWorkSheet.Cells(ROWNUM, 6) = DATE12
                    xlWorkSheet.Cells(ROWNUM, 7) = RO_DT
                    xlWorkSheet.Cells(ROWNUM, 8) = cugland
                    xlWorkSheet.Cells(ROWNUM, 9) = cugmob
                    xlWorkSheet.Cells(ROWNUM, 10) = TEXT14
                    xlWorkSheet.Cells(ROWNUM, 11) = DATE13
                    xlWorkSheet.Cells(ROWNUM, 12) = STFPOS
                    PREVDT = NPADT
                    PREVACNO = acno


                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("F" & ROWNUM, "F" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("G" & ROWNUM, "G" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("I" & ROWNUM, "I" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("J" & ROWNUM, "J" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("K" & ROWNUM, "K" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("L" & ROWNUM, "L" & ROWNUM).NumberFormat = "@"

                    xlWorkSheet.Range("A" & ROWNUM, "L" & ROWNUM).Font.Bold = True

                    acno = drx(0)
                    SCHEMECODE = drx(15)
                    DATE1 = drx(3)
                    NUMBER1 = drx(9)
                    NUMBER2 = drx(10)
                    DATE2 = drx(4)
                    NUMBER3 = drx(11)
                    NUMBER4 = drx(12)
                    TEXT1 = drx(14)
                    DATE4 = drx(5)
                    CRIAMTQTREND = CRIAMTQTREND = Math.Round(drx(28), 0, MidpointRounding.AwayFromZero)

                    ROWNUM = ROWNUM + 1

                    totnum = totnum + 1
                    totbal = totbal + NUMBER2
                    totcritic = totcritic + NUMBER4



                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = SCHEMECODE
                    xlWorkSheet.Cells(ROWNUM, 4) = DATE1
                    xlWorkSheet.Cells(ROWNUM, 5) = NUMBER1
                    xlWorkSheet.Cells(ROWNUM, 6) = NUMBER2
                    xlWorkSheet.Cells(ROWNUM, 7) = DATE2
                    xlWorkSheet.Cells(ROWNUM, 8) = NUMBER3
                    xlWorkSheet.Cells(ROWNUM, 9) = NUMBER4

                    CRITROW = ROWNUM
                    xlWorkSheet.Cells(ROWNUM, 10) = CRIAMTQTREND
                    xlWorkSheet.Cells(ROWNUM, 11) = TEXT1
                    xlWorkSheet.Cells(ROWNUM, 12) = DATE4

                    If NUMBER4 > 0 Then
                        xlWorkSheet.Range("I" & CRITROW).AddComment()
                        xlWorkSheet.Range("I" & CRITROW).Comment.Text("Balance to Crit.amt ratio=" & Format(NUMBER2 / NUMBER4, "fixed"))
                    End If

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("F" & ROWNUM, "F" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("G" & ROWNUM, "G" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("I" & ROWNUM, "I" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("J" & ROWNUM, "J" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("J" & ROWNUM, "K" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("K" & ROWNUM, "L" & ROWNUM).NumberFormat = "DD-MM-YYYY"

                Case 2  'S
                    acno = drx(0)
                    RePeriod = drx(9)
                    Installment = drx(10)
                    RepFrequency = drx(14)
                    FirstInstDate = drx(3)
                    WhetherRescheduled = drx(15)

                    ROWNUM = ROWNUM + 1

                    xlWorkSheet.Range("F" & ROWNUM, "G" & ROWNUM).Merge()
                    xlWorkSheet.Range("I" & ROWNUM, "L" & ROWNUM).Merge()

                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = "Rep Schedule"
                    xlWorkSheet.Cells(ROWNUM, 4) = RePeriod
                    xlWorkSheet.Cells(ROWNUM, 5) = Installment
                    xlWorkSheet.Cells(ROWNUM, 6) = RepFrequency
                    xlWorkSheet.Cells(ROWNUM, 8) = FirstInstDate
                    xlWorkSheet.Cells(ROWNUM, 9) = WhetherRescheduled

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("F" & ROWNUM, "F" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("I" & ROWNUM, "I" & ROWNUM).NumberFormat = "@"

                Case 3   'C
                    acno = drx(0)
                    CustID = drx(18)
                    CustomerName = drx(15)
                    Relation = drx(14)
                    GoldLoan = drx(10)
                    MobileNo = drx(16)
                    TotalDep = drx(9)

                    ROWNUM = ROWNUM + 1

                    xlWorkSheet.Range("E" & ROWNUM, "G" & ROWNUM).Merge()
                    xlWorkSheet.Range("H" & ROWNUM, "I" & ROWNUM).Merge()

                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = "Parties"
                    xlWorkSheet.Cells(ROWNUM, 4) = CustID
                    xlWorkSheet.Cells(ROWNUM, 5) = CustomerName
                    xlWorkSheet.Cells(ROWNUM, 8) = Relation
                    xlWorkSheet.Cells(ROWNUM, 10) = GoldLoan
                    xlWorkSheet.Cells(ROWNUM, 11) = MobileNo
                    xlWorkSheet.Cells(ROWNUM, 12) = TotalDep

                    If Relation = "A/C HOLDER" Then
                        If NUMBER4 > 0 Then
                            If TotalDep >= NUMBER4 Then
                                xlWorkSheet.Range("I" & CRITROW).Interior.Color = 13434825
                                xlWorkSheet.Range("I" & CRITROW).Font.Bold = True
                            End If
                        End If

                        If CRIAMTQTREND > 0 Then
                            If TotalDep >= CRIAMTQTREND Then
                                xlWorkSheet.Range("J" & CRITROW).Interior.Color = 13434825
                                xlWorkSheet.Range("J" & CRITROW).Font.Bold = True
                            End If
                        End If
                    End If

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("J" & ROWNUM, "J" & ROWNUM).NumberFormat = "##########"
                    xlWorkSheet.Range("K" & ROWNUM, "K" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("L" & ROWNUM, "L" & ROWNUM).NumberFormat = "##########"



                Case 4   'N
                    acno = drx(0)
                    NoticeDate = drx(3)
                    SendTo = drx(14)
                    NoticeName = drx(15)

                    ROWNUM = ROWNUM + 1

                    xlWorkSheet.Range("E" & ROWNUM, "G" & ROWNUM).Merge()
                    xlWorkSheet.Range("H" & ROWNUM, "L" & ROWNUM).Merge()

                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = "Notice"
                    xlWorkSheet.Cells(ROWNUM, 4) = NoticeDate
                    xlWorkSheet.Cells(ROWNUM, 5) = SendTo
                    xlWorkSheet.Cells(ROWNUM, 8) = NoticeName

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("K" & ROWNUM, "K" & ROWNUM).NumberFormat = "@"



                Case 5   'F
                    acno = drx(0)
                    FUDate = drx(3)
                    Contacted = drx(18)

                    If Contacted = "PARTY" Then Contacted = Contacted & "-" & drx(27)

                    ContType = drx(15)
                    InitiatedBy = drx(14)
                    DoneBy = drx(17)
                    Response = drx(19)


                    ROWNUM = ROWNUM + 1

                    xlWorkSheet.Range("I" & ROWNUM, "L" & ROWNUM).Merge()


                    xlWorkSheet.Cells(ROWNUM, 1) = NPADT
                    xlWorkSheet.Cells(ROWNUM, 2) = acno
                    xlWorkSheet.Cells(ROWNUM, 3) = "FollowUp"
                    xlWorkSheet.Cells(ROWNUM, 4) = FUDate
                    xlWorkSheet.Cells(ROWNUM, 5) = Contacted
                    xlWorkSheet.Cells(ROWNUM, 6) = ContType
                    xlWorkSheet.Cells(ROWNUM, 7) = InitiatedBy
                    xlWorkSheet.Cells(ROWNUM, 8) = DoneBy
                    xlWorkSheet.Cells(ROWNUM, 9) = Response

                    xlWorkSheet.Range("A" & ROWNUM, "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("B" & ROWNUM, "B" & ROWNUM).NumberFormat = "################"
                    xlWorkSheet.Range("C" & ROWNUM, "C" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("D" & ROWNUM, "D" & ROWNUM).NumberFormat = "DD-MM-YYYY"
                    xlWorkSheet.Range("E" & ROWNUM, "E" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("H" & ROWNUM, "H" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("J" & ROWNUM, "J" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("K" & ROWNUM, "K" & ROWNUM).NumberFormat = "@"
                    xlWorkSheet.Range("L" & ROWNUM, "L" & ROWNUM).NumberFormat = "@"


                Case Else

            End Select
        End While
        drx.Close()

        '=============================



        '=================================


        xlWorkSheet.Range("A9").Select()
        xlWorkSheet.Application.ActiveWindow.FreezePanes = True


        '================= SETTING FONT SIZE OF DATA ================================
        formatRange = xlWorkSheet.Range("A3", "L" & ROWNUM)
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"


        xlWorkSheet.Range("A3:B7").Merge()

        xlWorkSheet.Range("A3").HorizontalAlignment = 2
        xlWorkSheet.Range("A3").VerticalAlignment = 2
        xlWorkSheet.Range("A3").FormulaR1C1 = "No of Accounts : " & totnum & Chr(10) & "Bal O/s : " & Format(totbal / 100000, "fixed") & " Lakhs" & Chr(10) & "Crit Amt : " & Format(totcritic / 100000, "fixed") & " Lakhs"
        xlWorkSheet.Range("A3").Font.Bold = True
        xlWorkSheet.Range("A3").Font.Size = 11


        formatRange = xlWorkSheet.Range("A9", "L" & ROWNUM)
        formatRange.HorizontalAlignment = 2


        '==============PAGE & PRINT SETUP=================================

        With xlWorkSheet.PageSetup
            .PrintTitleRows = "$3:$8"
            .PrintTitleColumns = ""
        End With

        With xlWorkSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""

            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.05)
            .FooterMargin = xlApp.InchesToPoints(0.05)

            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With

        '================= END OF SHEET 2 ==========================================================================================================


        '================= BEGINNING OF SHEET 3 ==========================================================================================================



        xlNewSheet = xlWorkBook.Worksheets.Add(After:=xlWorkBook.Worksheets(2))

        xlNewSheet.Name = "Account Master"
        xlWorkSheet = xlNewSheet

        xlWorkSheet.Range("A1").ColumnWidth = 11
        xlWorkSheet.Range("B1").ColumnWidth = 15
        xlWorkSheet.Range("C1").ColumnWidth = 11
        xlWorkSheet.Range("D1").ColumnWidth = 11
        xlWorkSheet.Range("E1").ColumnWidth = 9
        xlWorkSheet.Range("F1").ColumnWidth = 9
        xlWorkSheet.Range("G1").ColumnWidth = 11
        xlWorkSheet.Range("H1").ColumnWidth = 10
        xlWorkSheet.Range("I1").ColumnWidth = 10
        xlWorkSheet.Range("J1").ColumnWidth = 10
        xlWorkSheet.Range("K1").ColumnWidth = 21
        xlWorkSheet.Range("L1").ColumnWidth = 11

        xlWorkSheet.Cells(1, 1) = "NPA dt."
        xlWorkSheet.Cells(1, 2) = "A/c No"
        xlWorkSheet.Cells(1, 3) = "Scheme"
        xlWorkSheet.Cells(1, 4) = "Open Date"
        xlWorkSheet.Cells(1, 5) = "Loan"
        xlWorkSheet.Cells(1, 6) = "Bal O/S"
        xlWorkSheet.Cells(1, 7) = "Due Date"
        xlWorkSheet.Cells(1, 8) = "Overdue"
        xlWorkSheet.Cells(1, 9) = "Critical Amt"
        xlWorkSheet.Cells(1, 10) = "Cri.Amt.Qtr.End"
        xlWorkSheet.Cells(1, 11) = "NPA Reason"
        xlWorkSheet.Cells(1, 12) = "AOD Due On"

        formatRange = xlWorkSheet.Range("A1", "L1")
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"
        formatRange.Font.Bold = True
        formatRange.HorizontalAlignment = 3


        Dim sql2 As String
        sql2 = "SELECT acno, TEXT2, DATE1, NUMBER1, NUMBER2,DATE2,NUMBER3, NUMBER4, TEXT1, DATE4, DATE3, NUMBER5 FROM C_MISADV WHERE NPAMAIN='M'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID =  " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO"
        Dim cmd2 As New OracleCommand(sql2, oracle_conn)
        Dim dr2 As OracleDataReader = cmd2.ExecuteReader()

        ROWNUM = 1
        While dr2.Read

            ROWNUM = ROWNUM + 1 ' Row no. of Excel

            '================= WRITING DATA TO EXCEL ================================

            processmessage(WORK_BOOK_NAME & ":Account Master : Writing row " & ROWNUM)
            Application.DoEvents()

            xlWorkSheet.Cells(ROWNUM, 1) = dr2(10)
            xlWorkSheet.Cells(ROWNUM, 2) = dr2(0)
            xlWorkSheet.Cells(ROWNUM, 3) = dr2(1)
            xlWorkSheet.Cells(ROWNUM, 4) = dr2(2)
            xlWorkSheet.Cells(ROWNUM, 5) = dr2(3)
            xlWorkSheet.Cells(ROWNUM, 6) = dr2(4)
            xlWorkSheet.Cells(ROWNUM, 7) = dr2(5)
            xlWorkSheet.Cells(ROWNUM, 8) = dr2(6)
            xlWorkSheet.Cells(ROWNUM, 9) = dr2(7)
            xlWorkSheet.Cells(ROWNUM, 10) = dr2(11)
            xlWorkSheet.Cells(ROWNUM, 11) = dr2(8)
            xlWorkSheet.Cells(ROWNUM, 12) = dr2(9)

        End While

        dr2.Close()

        xlWorkSheet.Range("A2", "A" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("B2", "B" & ROWNUM).NumberFormat = "################"
        xlWorkSheet.Range("C2", "C" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("D2", "D" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("E2", "E" & ROWNUM).NumberFormat = "##########"
        xlWorkSheet.Range("F2", "F" & ROWNUM).NumberFormat = "##########"
        xlWorkSheet.Range("G2", "G" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("H2", "H" & ROWNUM).NumberFormat = "##########"
        xlWorkSheet.Range("I2", "I" & ROWNUM).NumberFormat = "##########"
        xlWorkSheet.Range("I2", "J" & ROWNUM).NumberFormat = "##########"
        xlWorkSheet.Range("K2", "K" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("L2", "L" & ROWNUM).NumberFormat = "DD-MM-YYYY"

        xlWorkSheet.Range("A2").Select()
        xlWorkSheet.Application.ActiveWindow.FreezePanes = True

        '================= SETTING FONT SIZE OF DATA ================================
        formatRange = xlWorkSheet.Range("A2", "L" & ROWNUM)
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"

        '==============PAGE & PRINT SETUP=================================

        With xlWorkSheet.PageSetup
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = ""

        End With

        With xlWorkSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""

            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.05)
            .FooterMargin = xlApp.InchesToPoints(0.05)

            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With

        '================= END OF SHEET 3 ==========================================================================================================


        '================= BEGINNING OF SHEET 4 ==========================================================================================================


        xlNewSheet = xlWorkBook.Worksheets.Add(After:=xlWorkBook.Worksheets(3))

        xlNewSheet.Name = "Repayment Schedule"
        xlWorkSheet = xlNewSheet

        xlWorkSheet.Cells(1, 1) = "A/c No"
        xlWorkSheet.Cells(1, 2) = "Rep Period"
        xlWorkSheet.Cells(1, 3) = "Installment"
        xlWorkSheet.Cells(1, 4) = "Rep Frequency"
        xlWorkSheet.Cells(1, 5) = "First Inst Date"
        xlWorkSheet.Cells(1, 6) = "Whether Rescheduled In Finacle"

        xlWorkSheet.Range("A1").ColumnWidth = 14
        xlWorkSheet.Range("B1").ColumnWidth = 9
        xlWorkSheet.Range("C1").ColumnWidth = 9
        xlWorkSheet.Range("D1").ColumnWidth = 22
        xlWorkSheet.Range("E1").ColumnWidth = 10
        xlWorkSheet.Range("F1").ColumnWidth = 24

        formatRange = xlWorkSheet.Range("A1", "G1")
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"
        formatRange.Font.Bold = True
        formatRange.HorizontalAlignment = 3

        Dim sql3 As String
        sql3 = "select acno,NUMBER1,NUMBER2,TEXT1,DATE1, TEXT2 from  C_MISADV WHERE NPAMAIN='S'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID =  " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO"
        Dim cmd3 As New OracleCommand(sql3, oracle_conn)
        Dim dr3 As OracleDataReader = cmd3.ExecuteReader()

        ROWNUM = 1
        While dr3.Read

            ROWNUM = ROWNUM + 1 ' Row no. of Excel

            '================= WRITING DATA TO EXCEL ================================
            processmessage(WORK_BOOK_NAME & ":Rpayment Schedule : Writing row " & ROWNUM)
            Application.DoEvents()


            xlWorkSheet.Cells(ROWNUM, 1) = dr3(0)
            xlWorkSheet.Cells(ROWNUM, 2) = dr3(1)
            xlWorkSheet.Cells(ROWNUM, 3) = dr3(2)
            xlWorkSheet.Cells(ROWNUM, 4) = dr3(3)
            xlWorkSheet.Cells(ROWNUM, 5) = dr3(4)
            xlWorkSheet.Cells(ROWNUM, 6) = dr3(5)


        End While

        dr3.Close()

        xlWorkSheet.Range("A2", "A" & ROWNUM).NumberFormat = "################"
        xlWorkSheet.Range("B2", "B" & ROWNUM).NumberFormat = "##########"
        xlWorkSheet.Range("C2", "C" & ROWNUM).NumberFormat = "##########"
        xlWorkSheet.Range("D2", "D" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("E2", "E" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("F2", "F" & ROWNUM).NumberFormat = "@"

        xlWorkSheet.Range("A2").Select()
        xlWorkSheet.Application.ActiveWindow.FreezePanes = True


        '================= SETTING FONT SIZE OF DATA ================================
        formatRange = xlWorkSheet.Range("A2", "F" & ROWNUM)
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"


        '==============PAGE & PRINT SETUP=================================

        With xlWorkSheet.PageSetup
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = ""
        End With

        With xlWorkSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""

            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.05)
            .FooterMargin = xlApp.InchesToPoints(0.05)

            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With
        '================= END OF SHEET 4 ==========================================================================================================


        '================= BEGINNING OF SHEET 5 ==========================================================================================================


        xlNewSheet = xlWorkBook.Worksheets.Add(After:=xlWorkBook.Worksheets(4))

        xlNewSheet.Name = "Parties"
        xlWorkSheet = xlNewSheet

        xlWorkSheet.Cells(1, 1) = "A/c No"
        xlWorkSheet.Cells(1, 2) = "Cust ID"
        xlWorkSheet.Cells(1, 3) = "Customer Name"
        xlWorkSheet.Cells(1, 4) = "Relation"
        xlWorkSheet.Cells(1, 5) = "Gold Loan"
        xlWorkSheet.Cells(1, 6) = "Mobile No"
        xlWorkSheet.Cells(1, 7) = "Total Dep"

        xlWorkSheet.Range("A1").ColumnWidth = 14
        xlWorkSheet.Range("B1").ColumnWidth = 10
        xlWorkSheet.Range("C1").ColumnWidth = 32
        xlWorkSheet.Range("D1").ColumnWidth = 15
        xlWorkSheet.Range("E1").ColumnWidth = 9
        xlWorkSheet.Range("F1").ColumnWidth = 12
        xlWorkSheet.Range("G1").ColumnWidth = 10

        formatRange = xlWorkSheet.Range("A1", "G1")
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"
        formatRange.Font.Bold = True
        formatRange.HorizontalAlignment = 3

        Dim sql4 As String
        sql4 = "select acno, TEXT5,	TEXT2,TEXT1,NUMBER2,TEXT3,NUMBER1 from  C_MISADV WHERE NPAMAIN='C'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO, TEXT1"
        Dim cmd4 As New OracleCommand(sql4, oracle_conn)
        Dim dr4 As OracleDataReader = cmd4.ExecuteReader()

        ROWNUM = 1
        While dr4.Read

            ROWNUM = ROWNUM + 1 ' Row no. of Excel

            processmessage(WORK_BOOK_NAME & ":Parties : Writing row " & ROWNUM)
            Application.DoEvents()


            '================= WRITING DATA TO EXCEL ================================

            xlWorkSheet.Cells(ROWNUM, 1) = dr4(0)
            xlWorkSheet.Cells(ROWNUM, 2) = dr4(1)
            xlWorkSheet.Cells(ROWNUM, 3) = dr4(2)
            xlWorkSheet.Cells(ROWNUM, 4) = dr4(3)
            xlWorkSheet.Cells(ROWNUM, 5) = dr4(4)
            xlWorkSheet.Cells(ROWNUM, 6) = dr4(5)
            xlWorkSheet.Cells(ROWNUM, 7) = dr4(6)


        End While

        dr4.Close()

        xlWorkSheet.Range("A2", "A" & ROWNUM).NumberFormat = "################"
        xlWorkSheet.Range("B2", "B" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("C2", "C" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("D2", "D" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("E2", "E" & ROWNUM).NumberFormat = "##########"
        xlWorkSheet.Range("F2", "F" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("G2", "G" & ROWNUM).NumberFormat = "##########"

        xlWorkSheet.Range("A2").Select()
        xlWorkSheet.Application.ActiveWindow.FreezePanes = True

        '================= SETTING FONT SIZE OF DATA ================================
        formatRange = xlWorkSheet.Range("A2", "G" & ROWNUM)
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"


        '==============PAGE & PRINT SETUP=================================

        With xlWorkSheet.PageSetup
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = ""
        End With

        With xlWorkSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""

            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.05)
            .FooterMargin = xlApp.InchesToPoints(0.05)

            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With

        '================= END OF SHEET 5 ==========================================================================================================


        '================= BEGINNING OF SHEET 6 ==========================================================================================================


        xlNewSheet = xlWorkBook.Worksheets.Add(After:=xlWorkBook.Worksheets(5))

        xlNewSheet.Name = "Notice"
        xlWorkSheet = xlNewSheet


        xlWorkSheet.Cells(1, 1) = "A/c No"
        xlWorkSheet.Cells(1, 2) = "Notice Date"
        xlWorkSheet.Cells(1, 3) = "Send To"
        xlWorkSheet.Cells(1, 4) = "Notice Name"

        xlWorkSheet.Range("A1").ColumnWidth = 14
        xlWorkSheet.Range("B1").ColumnWidth = 10
        xlWorkSheet.Range("C1").ColumnWidth = 21
        xlWorkSheet.Range("D1").ColumnWidth = 32



        formatRange = xlWorkSheet.Range("A1", "E1")
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"
        formatRange.Font.Bold = True
        formatRange.HorizontalAlignment = 3

        Dim sql5 As String
        sql5 = "select acno, DATE1,TEXT1,TEXT2 from C_MISADV WHERE NPAMAIN='N'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO, DATE1"
        Dim cmd5 As New OracleCommand(sql5, oracle_conn)
        Dim dr5 As OracleDataReader = cmd5.ExecuteReader()

        ROWNUM = 1
        While dr5.Read

            ROWNUM = ROWNUM + 1 ' Row no. of Excel

            processmessage(WORK_BOOK_NAME & ":Notice : Writing row " & ROWNUM)
            Application.DoEvents()


            '================= WRITING DATA TO EXCEL ================================

            xlWorkSheet.Cells(ROWNUM, 1) = dr5(0)
            xlWorkSheet.Cells(ROWNUM, 2) = dr5(1)
            xlWorkSheet.Cells(ROWNUM, 3) = dr5(2)
            xlWorkSheet.Cells(ROWNUM, 4) = dr5(3)


        End While

        dr5.Close()

        xlWorkSheet.Range("A2", "A" & ROWNUM).NumberFormat = "################"
        xlWorkSheet.Range("B2", "B" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("C2", "C" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("D2", "D" & ROWNUM).NumberFormat = "@"


        xlWorkSheet.Range("A2").Select()
        xlWorkSheet.Application.ActiveWindow.FreezePanes = True


        '================= SETTING FONT SIZE OF DATA ================================
        formatRange = xlWorkSheet.Range("A2", "D" & ROWNUM)
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"


        '==============PAGE & PRINT SETUP=================================

        With xlWorkSheet.PageSetup
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = ""
        End With

        With xlWorkSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""

            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.05)
            .FooterMargin = xlApp.InchesToPoints(0.05)

            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With
        '================= END OF SHEET 6 ==========================================================================================================


        '================= BEGINNING OF SHEET 7 ==========================================================================================================


        xlNewSheet = xlWorkBook.Worksheets.Add(After:=xlWorkBook.Worksheets(6))

        xlNewSheet.Name = "Follow up"
        xlWorkSheet = xlNewSheet

        xlWorkSheet.Cells(1, 1) = "A/c No"
        xlWorkSheet.Cells(1, 2) = "Date"
        xlWorkSheet.Cells(1, 3) = "Contacted"
        xlWorkSheet.Cells(1, 4) = "Cont Type"
        xlWorkSheet.Cells(1, 5) = "Initiated By"
        xlWorkSheet.Cells(1, 6) = "Done By"
        xlWorkSheet.Cells(1, 7) = "Response"


        xlWorkSheet.Range("A1").ColumnWidth = 14
        xlWorkSheet.Range("B1").ColumnWidth = 9
        xlWorkSheet.Range("C1").ColumnWidth = 22
        xlWorkSheet.Range("D1").ColumnWidth = 15
        xlWorkSheet.Range("E1").ColumnWidth = 14
        xlWorkSheet.Range("F1").ColumnWidth = 20
        xlWorkSheet.Range("G1").ColumnWidth = 48

        formatRange = xlWorkSheet.Range("A1", "G1")
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"
        formatRange.Font.Bold = True
        formatRange.HorizontalAlignment = 3

        Dim sql6 As String
        sql6 = "select acno, DATE1,	TEXT5, TEXT6,TEXT2,TEXT1,TEXT3,TEXT4,TEXT7 from C_MISADV WHERE NPAMAIN='F'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO, DATE1"
        Dim cmd6 As New OracleCommand(sql6, oracle_conn)
        Dim dr6 As OracleDataReader = cmd6.ExecuteReader()

        ROWNUM = 1
        While dr6.Read

            ROWNUM = ROWNUM + 1 ' Row no. of Excel
            processmessage(WORK_BOOK_NAME & ":Followup : Writing row " & ROWNUM)
            Application.DoEvents()

            '================= WRITING DATA TO EXCEL ================================

            xlWorkSheet.Cells(ROWNUM, 1) = dr6(0)
            xlWorkSheet.Cells(ROWNUM, 2) = dr6(1)

            If dr6(2) = "PARTY" Then
                xlWorkSheet.Cells(ROWNUM, 3) = dr6(2) & "-" & dr6(3)
            Else
                xlWorkSheet.Cells(ROWNUM, 3) = dr6(2)
            End If

            xlWorkSheet.Cells(ROWNUM, 4) = dr6(4)
            xlWorkSheet.Cells(ROWNUM, 5) = dr6(5)
            xlWorkSheet.Cells(ROWNUM, 6) = dr6(6) & "-" & dr6(7)
            xlWorkSheet.Cells(ROWNUM, 7) = dr6(8)

        End While

        dr6.Close()

        xlWorkSheet.Range("A2", "A" & ROWNUM).NumberFormat = "################"
        xlWorkSheet.Range("B2", "B" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("C2", "C" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("D2", "D" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("E2", "E" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("F2", "F" & ROWNUM).NumberFormat = "@"

        xlWorkSheet.Range("A2").Select()
        xlWorkSheet.Application.ActiveWindow.FreezePanes = True



        '================= SETTING FONT SIZE OF DATA ================================
        formatRange = xlWorkSheet.Range("A2", "G" & ROWNUM)
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"

        '==============PAGE & PRINT SETUP=================================

        With xlWorkSheet.PageSetup
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = ""
        End With

        With xlWorkSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""

            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.05)
            .FooterMargin = xlApp.InchesToPoints(0.05)

            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With

        '================= END OF SHEET 7 ==========================================================================================================

        '================= BEGINNING OF SHEET 8 ==========================================================================================================


        xlNewSheet = xlWorkBook.Worksheets.Add(After:=xlWorkBook.Worksheets(7))

        xlNewSheet.Name = "Branch Details"
        xlWorkSheet = xlNewSheet

        xlWorkSheet.Cells(1, 1) = "SOLID"
        xlWorkSheet.Cells(1, 2) = "Branch Name"
        xlWorkSheet.Cells(1, 3) = "RO"
        xlWorkSheet.Cells(1, 4) = "Dist."
        xlWorkSheet.Cells(1, 5) = "OPEN DATE"

        xlWorkSheet.Cells(1, 6) = "ONLINE DATE"
        xlWorkSheet.Cells(1, 7) = "MANAGER"
        xlWorkSheet.Cells(2, 7) = "STAFF ID"
        xlWorkSheet.Cells(2, 8) = "Name"
        xlWorkSheet.Cells(2, 9) = "JOINED DATE"
        xlWorkSheet.Cells(1, 10) = "NO OF STAFF"

        xlWorkSheet.Cells(1, 11) = "No. of a/cs"
        xlWorkSheet.Cells(1, 12) = "Crit. Amt."
        xlWorkSheet.Cells(1, 13) = "Crit.Amt.Qtr.End"
        xlWorkSheet.Cells(1, 14) = "Overdue"
        xlWorkSheet.Cells(1, 15) = "Bal o/s"

        xlWorkSheet.Range("A1").ColumnWidth = 5
        xlWorkSheet.Range("B1").ColumnWidth = 30
        xlWorkSheet.Range("C1").ColumnWidth = 8
        xlWorkSheet.Range("D1").ColumnWidth = 8
        xlWorkSheet.Range("E1").ColumnWidth = 10
        xlWorkSheet.Range("F1").ColumnWidth = 10
        xlWorkSheet.Range("G1").ColumnWidth = 8
        xlWorkSheet.Range("H1").ColumnWidth = 25
        xlWorkSheet.Range("I1").ColumnWidth = 10
        xlWorkSheet.Range("J1").ColumnWidth = 10

        xlWorkSheet.Range("K1").ColumnWidth = 7
        xlWorkSheet.Range("L1").ColumnWidth = 7
        xlWorkSheet.Range("M1").ColumnWidth = 8
        xlWorkSheet.Range("N1").ColumnWidth = 8
        xlWorkSheet.Range("O1").ColumnWidth = 8

        formatRange = xlWorkSheet.Range("A1", "O2")
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"
        formatRange.Font.Bold = True
        formatRange.HorizontalAlignment = 3

        xlWorkSheet.Range("G1", "I1").Merge()

        xlWorkSheet.Range("A1", "A2").Merge()
        xlWorkSheet.Range("B1", "B2").Merge()
        xlWorkSheet.Range("C1", "C2").Merge()
        xlWorkSheet.Range("D1", "D2").Merge()
        xlWorkSheet.Range("E1", "E2").Merge()
        xlWorkSheet.Range("F1", "F2").Merge()
        xlWorkSheet.Range("J1", "J2").Merge()

        xlWorkSheet.Range("K1", "K2").Merge()
        xlWorkSheet.Range("L1", "L2").Merge()
        xlWorkSheet.Range("M1", "M2").Merge()
        xlWorkSheet.Range("N1", "N2").Merge()
        xlWorkSheet.Range("O1", "O2").Merge()

        xlWorkSheet.Range("K1").WrapText = True
        xlWorkSheet.Range("L1").WrapText = True
        xlWorkSheet.Range("M1").WrapText = True
        xlWorkSheet.Range("N1").WrapText = True
        xlWorkSheet.Range("O1").WrapText = True

        xlWorkSheet.Range("A1").VerticalAlignment = 3
        xlWorkSheet.Range("B1").VerticalAlignment = 3
        xlWorkSheet.Range("C1").VerticalAlignment = 3
        xlWorkSheet.Range("D1").VerticalAlignment = 3
        xlWorkSheet.Range("E1").VerticalAlignment = 3
        xlWorkSheet.Range("F1").VerticalAlignment = 3
        xlWorkSheet.Range("J1").VerticalAlignment = 3

        xlWorkSheet.Range("K1").VerticalAlignment = 3
        xlWorkSheet.Range("L1").VerticalAlignment = 3
        xlWorkSheet.Range("M1").VerticalAlignment = 3
        xlWorkSheet.Range("N1").VerticalAlignment = 3
        xlWorkSheet.Range("O1").VerticalAlignment = 3

        Dim SQLY As String = "update c_misadv A SET (NUMBER11,NUMBER12,NUMBER13,NUMBER14,NUMBER15) = (select COUNT(1), sum(number2), sum(number3), sum(number4), sum(number5) FROM C_MISADV B WHERE B.NPAMAIN='M' AND A.SOLID = B.SOLID AND B.NUMBER2 BETWEEN " & START_AMT & " AND " & END_AMT & " GROUP BY SOLID) WHERE A.NPAMAIN = 'Z' AND A.SOLID IN (SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")"

        oracle_execute_non_query("ten", username, username, SQLY)

        Dim sql7 As String
        sql7 = "SELECT solid,TEXT1,TEXT2,TEXT3,DATE1,DATE2,NUMBER1,text4,date3,NUMBER8, NUMBER11, NUMBER12, NUMBER13, NUMBER14, NUMBER15 from C_MISADV WHERE (NPAMAIN='Z') AND (NUMBER1 IS NOT NULL) AND (SOLID IN (SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY SOLID"
        Dim cmd7 As New OracleCommand(sql7, oracle_conn)
        Dim dr7 As OracleDataReader = cmd7.ExecuteReader()

        ROWNUM = 2
        While dr7.Read

            ROWNUM = ROWNUM + 1 ' Row no. of Excel
            processmessage(WORK_BOOK_NAME & ":Branch Details : Writing row " & ROWNUM)
            Application.DoEvents()

            '================= WRITING DATA TO EXCEL ================================

            xlWorkSheet.Cells(ROWNUM, 1) = dr7(0)
            xlWorkSheet.Cells(ROWNUM, 2) = UCase(dr7(1))
            xlWorkSheet.Cells(ROWNUM, 3) = dr7(2)
            xlWorkSheet.Cells(ROWNUM, 4) = dr7(3)
            xlWorkSheet.Cells(ROWNUM, 5) = dr7(4)
            xlWorkSheet.Cells(ROWNUM, 6) = dr7(5)
            xlWorkSheet.Cells(ROWNUM, 7) = dr7(6)
            xlWorkSheet.Cells(ROWNUM, 8) = dr7(7)
            xlWorkSheet.Cells(ROWNUM, 9) = dr7(8)
            xlWorkSheet.Cells(ROWNUM, 10) = dr7(9)

            xlWorkSheet.Cells(ROWNUM, 11) = dr7(10)
            xlWorkSheet.Cells(ROWNUM, 12) = dr7(13)
            xlWorkSheet.Cells(ROWNUM, 13) = dr7(14)
            xlWorkSheet.Cells(ROWNUM, 14) = dr7(12)
            xlWorkSheet.Cells(ROWNUM, 15) = dr7(11)

        End While

        dr7.Close()

        xlWorkSheet.Range("A3", "A" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("B3", "B" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("C3", "C" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("D3", "D" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("E3", "E" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("F3", "F" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("G3", "G" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("H3", "H" & ROWNUM).NumberFormat = "@"
        xlWorkSheet.Range("I3", "I" & ROWNUM).NumberFormat = "DD-MM-YYYY"
        xlWorkSheet.Range("J3", "J" & ROWNUM).NumberFormat = "###"

        xlWorkSheet.Range("K3", "K" & ROWNUM).NumberFormat = "###"
        xlWorkSheet.Range("L3", "L" & ROWNUM).NumberFormat = "###########"
        xlWorkSheet.Range("M3", "M" & ROWNUM).NumberFormat = "###########"
        xlWorkSheet.Range("N3", "N" & ROWNUM).NumberFormat = "###########"
        xlWorkSheet.Range("O3", "O" & ROWNUM).NumberFormat = "###########"



        xlWorkSheet.Range("A3").Select()
        xlWorkSheet.Application.ActiveWindow.FreezePanes = True

        '================= SETTING FONT SIZE OF DATA ================================
        formatRange = xlWorkSheet.Range("A3", "O" & ROWNUM)
        formatRange.Font.Size = 9
        formatRange.Font.Name = "Arial"

        '==============PAGE & PRINT SETUP=================================

        With xlWorkSheet.PageSetup
            .PrintTitleRows = "$1:$2"
            .PrintTitleColumns = ""
        End With

        With xlWorkSheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = ""

            .LeftMargin = xlApp.InchesToPoints(0.5)
            .RightMargin = xlApp.InchesToPoints(0.5)
            .TopMargin = xlApp.InchesToPoints(0.5)
            .BottomMargin = xlApp.InchesToPoints(0.5)
            .HeaderMargin = xlApp.InchesToPoints(0.05)
            .FooterMargin = xlApp.InchesToPoints(0.05)

            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = True
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With

        '================= END OF SHEET 8 ==========================================================================================================

        '================= NAMING, SAVING & CLOSING WORKBOOK ================================

        PREVDT = ""
        PREVACNO = ""

        xlWorkBook.Sheets("All In One - Date Wise").Select()


        xlWorkBook.SaveAs("D:\PNPA\" & WORK_BOOK_NAME, 51, misValue, misValue, misValue, misValue, _
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)

        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()

        '================= RELEASING EXCEL ================================

        releaseEXCELObject(xlWorkSheet)
        releaseEXCELObject(xlWorkBook)
        releaseEXCELObject(xlApp)

        processmessage(WORK_BOOK_NAME & ":Over.")
        Application.DoEvents()
        oracle_conn.Close()

    End Sub

    Public Sub prnt4excel(ByVal sw As StreamWriter, ByVal typ As String, ByVal param6 As String, Optional ByVal sheetno As Integer = 1, Optional ByVal rowno1 As Long = 1, Optional ByVal colno1 As Integer = 1, Optional ByVal rowno2 As Long = 1, Optional ByVal colno2 As Integer = 1)
        sw.WriteLine(UCase(typ) & "~" & param6 & "~" & sheetno & "~" & rowno1 & "~" & colno1 & "~" & rowno2 & "~" & colno2)
    End Sub

    Private Sub GEN_NPA_XL_MACRO(ByVal SOLSET As String, ByVal START_AMT As Double, ByVal END_AMT As Double, ByVal FILENAME As String)  'NPA FOLLOWUP XL DATA GENERATION
        Dim WORK_BOOK_NAME As String

        If START_AMT > 0 Then
            WORK_BOOK_NAME = "d:\pnpa\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_" & FILENAME & "_" & Trim(START_AMT) & "_" & Trim(END_AMT) & ".xlsx"
        Else
            WORK_BOOK_NAME = "d:\pnpa\7DAY_NPA_THREAT_" & Format(RptDate, "ddMMyyyy") & "_" & FILENAME & "_" & "Upto" & "_" & Trim(END_AMT) & ".xlsx"
        End If

        SOLSET = "'" & SOLSET & "'"

        Dim txtfilename As String
        txtfilename = Replace(WORK_BOOK_NAME, "xlsx", "txt")

        '==============================================

        Dim oracle_cnn_string As String = "Data Source=ten; User Id= " & username & ";Password= " & username & ";"
        Dim oracle_conn As New OracleConnection(oracle_cnn_string)
        oracle_conn.Open()

        '================= CREATING EXCEL FILE ================================
        Dim sw As StreamWriter = New StreamWriter(txtfilename)

        Call prnt4excel(sw, "FILENAME", WORK_BOOK_NAME)
        Call prnt4excel(sw, "totsheets", 8)

        Dim sql1 As String

        sql1 = "SELECT ACNO,SCHEMECODE,SOLID,NVL(DATE1,'01-JAN-1901'),NVL(DATE2,'01-JAN-1901'),NVL(DATE4,'01-JAN-1901'),NVL(DATE11,'01-JAN-1901'),NVL(DATE12,'01-JAN-1901'),NVL(DATE13,'01-JAN-1901'),NVL(NUMBER1,0),NVL(NUMBER2,0),NVL(NUMBER3,0),NVL(NUMBER4,0),NVL(NUMBER18,0),NVL(TEXT1,' '),NVL(TEXT2,' '),NVL(TEXT3,' '),NVL(TEXT4,' '),NVL(TEXT5,' '),NVL(TEXT7,' '),NVL(TEXT11,' '),NVL(TEXT12,' '),NVL(TEXT13,' '),NVL(TEXT14,' '), NUMBER20, DATE20, NUMBER21, NVL(TEXT6,' '), NVL(NUMBER5,0) FROM C_MISADV WHERE ( NUMBER20>0 ) AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO, NUMBER20, TEXT1, DATE1"
        Dim cmd1 As New OracleCommand(sql1, oracle_conn)
        Dim dr1 As OracleDataReader = cmd1.ExecuteReader()

        '================= BEGINNING OF SHEET 1 ==========================================================================================================

        '================= NAMING SHEET 1 ================================
        Call prnt4excel(sw, "shtname", "All In One - Date Wise", 1)

        Dim PREVDT As String = ""
        Dim PREVACNO As String = ""

        '================= SETTING COLUMN WIDTHS ================================

        Call prnt4excel(sw, "colw", 9, 1, , 1)
        Call prnt4excel(sw, "colw", 13, 1, , 2)
        Call prnt4excel(sw, "colw", 11, 1, , 3)
        Call prnt4excel(sw, "colw", 18, 1, , 4)
        Call prnt4excel(sw, "colw", 9, 1, , 5)
        Call prnt4excel(sw, "colw", 9, 1, , 6)
        Call prnt4excel(sw, "colw", 13, 1, , 7)
        Call prnt4excel(sw, "colw", 11, 1, , 8)
        Call prnt4excel(sw, "colw", 10, 1, , 9)
        Call prnt4excel(sw, "colw", 22, 1, , 10)
        Call prnt4excel(sw, "colw", 10, 1, , 11)
        Call prnt4excel(sw, "colw", 10, 1, , 12)

        '================= Setting Column Headings ================================

        Call prnt4excel(sw, "txt", "NPA THREAT FOR THE NEXT 7 DAYS AS ON " & Format(RptDate, "dd-MM-yyyy"), 1, 1, 1)
        Call prnt4excel(sw, "txt", "SOLSET : " & Replace(SOLSET, "'", "") & " (OUTSTANDING BETWEEN " & Trim(START_AMT) & " AND " & Trim(END_AMT) & ")", 1, 2, 1)
        Call prnt4excel(sw, "merge", "", 1, 1, 1, 1, 12)
        Call prnt4excel(sw, "hal", 3, 1, 1, 1, 1, 12)
        Call prnt4excel(sw, "fntnam", "Arial", 1, 1, 1, 1, 12)
        Call prnt4excel(sw, "FNTSIZ", 11, 1, 1, 1, 1, 12)
        Call prnt4excel(sw, "fntbold", "", 1, 1, 1, 1, 12)
        Call prnt4excel(sw, "fntnam", "Arial", 1, 1, 1, 1, 12)
        Call prnt4excel(sw, "merge", "", 1, 2, 1, 2, 12)
        Call prnt4excel(sw, "hal", 3, 1, 2, 1, 2, 12)
        Call prnt4excel(sw, "fntnam", "", 1, 2, 1, 2, 12)
        Call prnt4excel(sw, "FNTSIZ", 10, 1, 2, 1, 2, 12)
        Call prnt4excel(sw, "fntbold", "", 1, 2, 1, 2, 12)
        Call prnt4excel(sw, "txt", "Branch", 1, 3, 3)
        Call prnt4excel(sw, "txt", "Branch Name", 1, 3, 4)
        Call prnt4excel(sw, "txt", "Open Date", 1, 3, 5)
        Call prnt4excel(sw, "txt", "Online Date", 1, 3, 6)
        Call prnt4excel(sw, "txt", "RO/DT", 1, 3, 7)
        Call prnt4excel(sw, "txt", "CUG - Office", 1, 3, 8)
        Call prnt4excel(sw, "txt", "CUG - Mobile", 1, 3, 9)
        Call prnt4excel(sw, "txt", "Manager", 1, 3, 10)
        Call prnt4excel(sw, "txt", "Since", 1, 3, 11)
        Call prnt4excel(sw, "txt", "Staff Strength", 1, 3, 12)
        Call prnt4excel(sw, "txt", "Scheme", 1, 4, 3)
        Call prnt4excel(sw, "txt", "Open Date", 1, 4, 4)
        Call prnt4excel(sw, "txt", "Loan", 1, 4, 5)
        Call prnt4excel(sw, "txt", "Bal O/S", 1, 4, 6)
        Call prnt4excel(sw, "txt", "Due Date", 1, 4, 7)
        Call prnt4excel(sw, "txt", "Overdue", 1, 4, 8)
        Call prnt4excel(sw, "txt", "Critical Amt", 1, 4, 9)
        Call prnt4excel(sw, "txt", "Cri.Amt.Qtr.End", 1, 4, 10)
        Call prnt4excel(sw, "txt", "NPA Reason", 1, 4, 11)
        Call prnt4excel(sw, "txt", "AOD Due On", 1, 4, 12)
        Call prnt4excel(sw, "txt", "Rep Schedule", 1, 5, 3)
        Call prnt4excel(sw, "txt", "Rep Period", 1, 5, 4)
        Call prnt4excel(sw, "txt", "Installment", 1, 5, 5)
        Call prnt4excel(sw, "txt", "Rep Frequency", 1, 5, 6)
        Call prnt4excel(sw, "txt", "First Inst Date", 1, 5, 8)
        Call prnt4excel(sw, "txt", "Whether Rescheduled In Finacle", 1, 5, 9)
        Call prnt4excel(sw, "txt", "Parties", 1, 6, 3)
        Call prnt4excel(sw, "txt", "Cust ID", 1, 6, 4)
        Call prnt4excel(sw, "txt", "Customer Name", 1, 6, 5)
        Call prnt4excel(sw, "txt", "Relation", 1, 6, 8)
        Call prnt4excel(sw, "txt", "Gold Loan", 1, 6, 10)
        Call prnt4excel(sw, "txt", "Mobile No", 1, 6, 11)
        Call prnt4excel(sw, "txt", "TotalDep", 1, 6, 12)
        Call prnt4excel(sw, "txt", "Notice", 1, 7, 3)
        Call prnt4excel(sw, "txt", "Notice Date", 1, 7, 4)
        Call prnt4excel(sw, "txt", "Send To", 1, 7, 5)
        Call prnt4excel(sw, "txt", "Notice Name", 1, 7, 8)
        Call prnt4excel(sw, "txt", "NPA dt.", 1, 8, 1)
        Call prnt4excel(sw, "txt", "A/c No", 1, 8, 2)
        Call prnt4excel(sw, "txt", "Follow Up", 1, 8, 3)
        Call prnt4excel(sw, "txt", "Date", 1, 8, 4)
        Call prnt4excel(sw, "txt", "Contacted", 1, 8, 5)
        Call prnt4excel(sw, "txt", "Cont Type", 1, 8, 6)
        Call prnt4excel(sw, "txt", "Initiated By", 1, 8, 7)
        Call prnt4excel(sw, "txt", "Done By", 1, 8, 8)
        Call prnt4excel(sw, "txt", "Response", 1, 8, 9)
        Call prnt4excel(sw, "HAL", 3, 1, 3, 1, 12, 8)
        Call prnt4excel(sw, "FNTBOLD", "", 1, 3, 1, 8, 12)
        Call prnt4excel(sw, "FNTSIZ", 9, 1, 3, 1, 12, 8)
        Call prnt4excel(sw, "MERGE", "", 1, 5, 6, 5, 7)
        Call prnt4excel(sw, "MERGE", "", 1, 5, 9, 5, 12)
        Call prnt4excel(sw, "MERGE", "", 1, 6, 5, 6, 7)
        Call prnt4excel(sw, "MERGE", "", 1, 6, 8, 6, 9)
        Call prnt4excel(sw, "MERGE", "", 1, 7, 5, 7, 7)
        Call prnt4excel(sw, "MERGE", "", 1, 7, 8, 7, 12)
        Call prnt4excel(sw, "MERGE", "", 1, 8, 9, 8, 12)

        Call prnt4excel(sw, "HAL", 3, 1, 3, 9, 8, 12)
        Call prnt4excel(sw, "FNTBOLD", 3, 1, 3, 9, 8, 12)




        Dim CRITROW As Long
        Dim npamainnum As Integer
        Dim acno As String
        Dim Contacted As String
        Dim ContType As String
        Dim cugland As String
        Dim cugmob As String
        Dim CustID As String
        Dim CustomerName As String
        Dim DATE1 As Date
        Dim DATE11 As Date
        Dim DATE12 As Date
        Dim DATE13 As Date
        Dim DATE2 As Date
        Dim DATE4 As Date
        Dim DoneBy As String
        Dim FirstInstDate As Date
        Dim FUDate As Date
        Dim GoldLoan As String
        Dim InitiatedBy As String
        Dim Installment As String
        Dim MobileNo As String
        Dim NoticeDate As Date
        Dim NoticeName As String
        Dim NUMBER1 As Double
        Dim NUMBER2 As Double
        Dim NUMBER3 As Double
        Dim NUMBER4 As Double
        Dim Relation As String
        Dim RePeriod As Double
        Dim RepFrequency As String
        Dim Response As String
        Dim RO_DT As String
        Dim SCHEMECODE As String
        Dim SendTo As String
        Dim solid As String
        Dim STFPOS As Integer
        Dim TEXT1 As String
        Dim TEXT11 As String
        Dim TEXT14 As String
        Dim TotalDep As Double
        Dim WhetherRescheduled As String
        Dim NPADT As Date
        Dim CRIAMTQTREND As Double

        Dim ROWNUM As Integer
        ROWNUM = 8

        '================= WRITING DATA TO EXCEL ================================
        Dim totnum As Integer = 0
        Dim totbal As Double = 0
        Dim totcritic As Double = 0

        While dr1.Read
            npamainnum = dr1(24)
            Select Case npamainnum
                Case 1
                    acno = dr1(0)
                    solid = dr1(2)
                    TEXT11 = dr1(20)
                    DATE11 = dr1(6)
                    DATE12 = dr1(7)
                    RO_DT = dr1(21) & "/" & dr1(22)
                    cugland = Trim(Val(solid) - 40000 + 4000)
                    cugmob = Trim(Val(solid) - 40000 + 9400999000)
                    TEXT14 = dr1(23)
                    DATE13 = dr1(8)
                    STFPOS = dr1(13)
                    NPADT = dr1(25)

                    ROWNUM = ROWNUM + 1

                    If ROWNUM <> 9 Then
                        prnt4excel(sw, "TXT", PREVDT, 1, ROWNUM, 1)
                        prnt4excel(sw, "TXT", PREVACNO, 1, ROWNUM, 2)
                        prnt4excel(sw, "nf", "################", 1, ROWNUM, 2, ROWNUM, 2)
                        prnt4excel(sw, "TXT", "Remarks", 1, ROWNUM, 3)
                        Call prnt4excel(sw, "MERGE", "", 1, ROWNUM, 4, ROWNUM, 12)
                        ROWNUM = ROWNUM + 1
                    End If

                    processmessage(Mid(WORK_BOOK_NAME, 9) & ":All in one : Writing row " & ROWNUM)
                    Application.DoEvents()
                    prnt4excel(sw, "TXT", NPADT, 1, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 1, ROWNUM, 2)
                    prnt4excel(sw, "TXT", solid, 1, ROWNUM, 3)
                    prnt4excel(sw, "TXT", UCase(TEXT11), 1, ROWNUM, 4)
                    prnt4excel(sw, "TXT", DATE11, 1, ROWNUM, 5)
                    prnt4excel(sw, "TXT", DATE12, 1, ROWNUM, 6)
                    prnt4excel(sw, "TXT", RO_DT, 1, ROWNUM, 7)
                    prnt4excel(sw, "TXT", cugland, 1, ROWNUM, 8)
                    prnt4excel(sw, "TXT", cugmob, 1, ROWNUM, 9)
                    prnt4excel(sw, "TXT", TEXT14, 1, ROWNUM, 10)
                    prnt4excel(sw, "TXT", DATE13, 1, ROWNUM, 11)
                    prnt4excel(sw, "TXT", STFPOS, 1, ROWNUM, 12)

                    PREVDT = NPADT
                    PREVACNO = acno

                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 1, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 6, ROWNUM, 6)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 7, ROWNUM, 7)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 9, ROWNUM, 9)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 10, ROWNUM, 10)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 11, ROWNUM, 11)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 12, ROWNUM, 12)
                    prnt4excel(sw, "FNTBOLD", "", 1, ROWNUM, 1, ROWNUM, 12)

                    acno = dr1(0)
                    SCHEMECODE = dr1(15)
                    DATE1 = dr1(3)
                    NUMBER1 = dr1(9)
                    NUMBER2 = dr1(10)
                    DATE2 = dr1(4)
                    NUMBER3 = dr1(11)
                    NUMBER4 = dr1(12)
                    TEXT1 = dr1(14)
                    DATE4 = dr1(5)
                    CRIAMTQTREND = Math.Round(dr1(28), 0, MidpointRounding.AwayFromZero)

                    ROWNUM = ROWNUM + 1

                    totnum = totnum + 1
                    totbal = totbal + NUMBER2
                    totcritic = totcritic + NUMBER4

                    prnt4excel(sw, "TXT", NPADT, 1, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 1, ROWNUM, 2)
                    prnt4excel(sw, "TXT", SCHEMECODE, 1, ROWNUM, 3)
                    prnt4excel(sw, "TXT", DATE1, 1, ROWNUM, 4)
                    prnt4excel(sw, "TXT", NUMBER1, 1, ROWNUM, 5)
                    prnt4excel(sw, "TXT", NUMBER2, 1, ROWNUM, 6)
                    prnt4excel(sw, "TXT", DATE2, 1, ROWNUM, 7)
                    prnt4excel(sw, "TXT", NUMBER3, 1, ROWNUM, 8)
                    prnt4excel(sw, "TXT", NUMBER4, 1, ROWNUM, 9)

                    CRITROW = ROWNUM
                    prnt4excel(sw, "TXT", CRIAMTQTREND, 1, ROWNUM, 10)
                    prnt4excel(sw, "TXT", TEXT1, 1, ROWNUM, 11)
                    prnt4excel(sw, "TXT", DATE4, 1, ROWNUM, 12)

                    If NUMBER4 > 0 Then
                        Call prnt4excel(sw, "COMM", "Balance to Crit.amt ratio=" & Format(NUMBER2 / NUMBER4, "fixed"), 1, CRITROW, 9)
                    End If

                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 1, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 6, ROWNUM, 6)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 7, ROWNUM, 7)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 9, ROWNUM, 9)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 10, ROWNUM, 10)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 11, ROWNUM, 11)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 12, ROWNUM, 12)

                Case 2  'S
                    acno = dr1(0)
                    RePeriod = dr1(9)
                    Installment = dr1(10)
                    RepFrequency = dr1(14)
                    FirstInstDate = dr1(3)
                    WhetherRescheduled = dr1(15)

                    ROWNUM = ROWNUM + 1

                    prnt4excel(sw, "MERGE", "", 1, ROWNUM, 6, ROWNUM, 7)
                    prnt4excel(sw, "MERGE", "", 1, ROWNUM, 8, ROWNUM, 12)
                    prnt4excel(sw, "TXT", NPADT, 1, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 1, ROWNUM, 2)
                    prnt4excel(sw, "TXT", "Rep Schedule", 1, ROWNUM, 3)
                    prnt4excel(sw, "TXT", RePeriod, 1, ROWNUM, 4)
                    prnt4excel(sw, "TXT", Installment, 1, ROWNUM, 5)
                    prnt4excel(sw, "TXT", RepFrequency, 1, ROWNUM, 6)
                    prnt4excel(sw, "TXT", FirstInstDate, 1, ROWNUM, 8)
                    prnt4excel(sw, "TXT", WhetherRescheduled, 1, ROWNUM, 9)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 1, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 6, ROWNUM, 6)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 7, ROWNUM, 7)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 8, ROWNUM, 8)

                Case 3   'C
                    acno = dr1(0)
                    CustID = dr1(18)
                    CustomerName = dr1(15)
                    Relation = dr1(14)
                    GoldLoan = dr1(10)
                    MobileNo = dr1(16)
                    TotalDep = dr1(9)

                    ROWNUM = ROWNUM + 1

                    prnt4excel(sw, "MERGE", "", 1, ROWNUM, 5, ROWNUM, 7)
                    prnt4excel(sw, "MERGE", "", 1, ROWNUM, 8, ROWNUM, 9)
                    prnt4excel(sw, "TXT", NPADT, 1, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 1, ROWNUM, 2)
                    prnt4excel(sw, "TXT", "Parties", 1, ROWNUM, 3)
                    prnt4excel(sw, "TXT", CustID, 1, ROWNUM, 4)
                    prnt4excel(sw, "TXT", CustomerName, 1, ROWNUM, 5)
                    prnt4excel(sw, "TXT", Relation, 1, ROWNUM, 8)
                    prnt4excel(sw, "TXT", GoldLoan, 1, ROWNUM, 10)
                    prnt4excel(sw, "TXT", MobileNo, 1, ROWNUM, 11)
                    prnt4excel(sw, "TXT", TotalDep, 1, ROWNUM, 12)

                    If Relation = "A/C HOLDER" Then
                        If NUMBER4 > 0 Then
                            If TotalDep >= NUMBER4 Then
                                prnt4excel(sw, "BG", 13434825, 1, CRITROW, 9, CRITROW, 9)
                                prnt4excel(sw, "FNTBOLD", "", 1, CRITROW, 9, CRITROW, 9)
                            End If
                        End If

                        If CRIAMTQTREND > 0 Then
                            If TotalDep >= CRIAMTQTREND Then
                                prnt4excel(sw, "BG", 13434825, 1, CRITROW, 10, CRITROW, 10)
                                prnt4excel(sw, "FNTBOLD", "", 1, CRITROW, 10, CRITROW, 10)
                            End If
                        End If
                    End If

                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 1, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 10, ROWNUM, 10)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 11, ROWNUM, 11)
                    prnt4excel(sw, "nf", "##########", 1, ROWNUM, 12, ROWNUM, 12)

                Case 4   'N
                    acno = dr1(0)
                    NoticeDate = dr1(3)
                    SendTo = dr1(14)
                    NoticeName = dr1(15)

                    ROWNUM = ROWNUM + 1

                    prnt4excel(sw, "MERGE", "", 1, ROWNUM, 5, ROWNUM, 7)
                    prnt4excel(sw, "MERGE", "", 1, ROWNUM, 8, ROWNUM, 12)
                    prnt4excel(sw, "TXT", NPADT, 1, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 1, ROWNUM, 2)
                    prnt4excel(sw, "TXT", "Notice", 1, ROWNUM, 3)
                    prnt4excel(sw, "TXT", NoticeDate, 1, ROWNUM, 4)
                    prnt4excel(sw, "TXT", SendTo, 1, ROWNUM, 5)
                    prnt4excel(sw, "TXT", NoticeName, 1, ROWNUM, 8)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 1, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 11, ROWNUM, 11)

                Case 5   'F
                    acno = dr1(0)
                    FUDate = dr1(3)
                    Contacted = dr1(18)

                    If Contacted = "PARTY" Then Contacted = Contacted & "-" & dr1(27)

                    ContType = dr1(15)
                    InitiatedBy = dr1(14)
                    DoneBy = dr1(17)
                    Response = dr1(19)

                    ROWNUM = ROWNUM + 1

                    prnt4excel(sw, "MERGE", "", 1, ROWNUM, 9, ROWNUM, 12)
                    prnt4excel(sw, "TXT", NPADT, 1, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 1, ROWNUM, 2)
                    prnt4excel(sw, "TXT", "FollowUp", 1, ROWNUM, 3)
                    prnt4excel(sw, "TXT", FUDate, 1, ROWNUM, 4)
                    prnt4excel(sw, "TXT", Contacted, 1, ROWNUM, 5)
                    prnt4excel(sw, "TXT", ContType, 1, ROWNUM, 6)
                    prnt4excel(sw, "TXT", InitiatedBy, 1, ROWNUM, 7)
                    prnt4excel(sw, "TXT", DoneBy, 1, ROWNUM, 8)
                    prnt4excel(sw, "TXT", Response, 1, ROWNUM, 9)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 1, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 1, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 10, ROWNUM, 10)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 11, ROWNUM, 11)
                    prnt4excel(sw, "nf", "@", 1, ROWNUM, 12, ROWNUM, 12)

                Case Else
            End Select
        End While
        dr1.Close()

        '================= SETTING FNT SIZE OF DATA ================================

        prnt4excel(sw, "FNTSIZ", 9, 1, 3, 1, ROWNUM, 12)
        prnt4excel(sw, "FNTNAM", "Arial", 1, 3, 1, ROWNUM, 12)
        prnt4excel(sw, "MERGE", "", 1, 3, 1, 7, 2)
        prnt4excel(sw, "hal", 2, 1, 3, 1, 3, 1)
        prnt4excel(sw, "val", 2, 1, 3, 1, 3, 1)
        prnt4excel(sw, "txt", "No of Accounts : " & totnum & Chr(10) & "Bal O/s : " & Format(totbal / 100000, "fixed") & " Lakhs" & Chr(10) & "Crit Amt : " & Format(totcritic / 100000, "fixed") & " Lakhs", 1, 3, 1, 3, 1)
        prnt4excel(sw, "fntbold", "", 1, 3, 1, 3, 1)
        prnt4excel(sw, "FNTSIZ", 11, 1, 3, 1, 3, 1)
        prnt4excel(sw, "hal", 2, 1, 9, 1, ROWNUM, 12)
        prnt4excel(sw, "FP", "", 1, 9, 1)

        '================= END OF SHEET 1 ==========================================================================================================

        Dim SQLX As String = "SELECT ACNO,SCHEMECODE,SOLID,NVL(DATE1,'01-JAN-1901'),NVL(DATE2,'01-JAN-1901'),NVL(DATE4,'01-JAN-1901'),NVL(DATE11,'01-JAN-1901'),NVL(DATE12,'01-JAN-1901'),NVL(DATE13,'01-JAN-1901'),NVL(NUMBER1,0),NVL(NUMBER2,0),NVL(NUMBER3,0),NVL(NUMBER4,0),NVL(NUMBER18,0),NVL(TEXT1,' '),NVL(TEXT2,' '),NVL(TEXT3,' '),NVL(TEXT4,' '),NVL(TEXT5,' '),NVL(TEXT7,' '),NVL(TEXT11,' '),NVL(TEXT12,' '),NVL(TEXT13,' '),NVL(TEXT14,' '), NUMBER20, DATE20, NUMBER21, NVL(TEXT6,' '), NVL(NUMBER5,0) FROM C_MISADV WHERE ( NUMBER20>0 ) AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY SOLID, DATE20, NUMBER21 DESC, ACNO, NUMBER20, TEXT1, DATE1"
        Dim cmdx As New OracleCommand(SQLX, oracle_conn)
        Dim drx As OracleDataReader = cmdx.ExecuteReader()

        '================= BEGINNING OF SHEET 2 ==========================================================================================================

        '================= NAMING SHEET 2 ================================

        prnt4excel(sw, "shtname", "All In One - SOL Wise", 2)

        '================= SETTING COLUMN WIDTHS ================================

        prnt4excel(sw, "colw", 9, 2, , 1)
        prnt4excel(sw, "colw", 13, 2, , 2)
        prnt4excel(sw, "colw", 11, 2, , 3)
        prnt4excel(sw, "colw", 18, 2, , 4)
        prnt4excel(sw, "colw", 9, 2, , 5)
        prnt4excel(sw, "colw", 9, 2, , 6)
        prnt4excel(sw, "colw", 13, 2, , 7)
        prnt4excel(sw, "colw", 11, 2, , 8)
        prnt4excel(sw, "colw", 10, 2, , 9)
        prnt4excel(sw, "colw", 22, 2, , 10)
        prnt4excel(sw, "colw", 10, 2, , 11)

        '================= Setting Column Headings ================================

        prnt4excel(sw, "txt", "NPA THREAT FOR THE NEXT 7 DAYS AS ON " & Format(RptDate, "dd-MM-yyyy"), 2, 1, 1)
        prnt4excel(sw, "txt", "SOLSET : " & Replace(SOLSET, "'", "") & " (OUTSTANDING BETWEEN " & Trim(START_AMT) & " AND " & Trim(END_AMT) & ")", 2, 2, 1)
        prnt4excel(sw, "MERGE", "", 2, 1, 1, 1, 12)
        prnt4excel(sw, "hal", 3, 2, 1, 1, 1, 12)
        prnt4excel(sw, "fntnam", "Arial", 2, 1, 1, 1, 12)
        prnt4excel(sw, "FNTSIZ", 11, 2, 1, 1, 1, 12)
        prnt4excel(sw, "fntbold", "", 2, 1, 1, 1, 12)
        prnt4excel(sw, "MERGE", "", 2, 2, 1, 2, 12)
        prnt4excel(sw, "hal", 3, 2, 2, 1, 2, 12)
        prnt4excel(sw, "fntnam", "Arial", 2, 2, 1, 2, 12)
        prnt4excel(sw, "FNTSIZ", 11, 2, 2, 1, 2, 12)
        prnt4excel(sw, "fntbold", "", 2, 2, 1, 2, 12)
        prnt4excel(sw, "TXT", "Branch", 2, 3, 3)
        prnt4excel(sw, "TXT", "Branch Name", 2, 3, 4)
        prnt4excel(sw, "TXT", "Open Date", 2, 3, 5)
        prnt4excel(sw, "TXT", "Online Date", 2, 3, 6)
        prnt4excel(sw, "TXT", "RO/DT", 2, 3, 7)
        prnt4excel(sw, "TXT", "CUG - Office", 2, 3, 8)
        prnt4excel(sw, "TXT", "CUG - Mobile", 2, 3, 9)
        prnt4excel(sw, "TXT", "Manager", 2, 3, 10)
        prnt4excel(sw, "TXT", "Since", 2, 3, 11)
        prnt4excel(sw, "TXT", "Staff Strength", 2, 3, 12)
        prnt4excel(sw, "TXT", "Scheme", 2, 4, 3)
        prnt4excel(sw, "TXT", "Open Date", 2, 4, 4)
        prnt4excel(sw, "TXT", "Loan", 2, 4, 5)
        prnt4excel(sw, "TXT", "Bal O/S", 2, 4, 6)
        prnt4excel(sw, "TXT", "Due Date", 2, 4, 7)
        prnt4excel(sw, "TXT", "Overdue", 2, 4, 8)
        prnt4excel(sw, "TXT", "Critical Amt", 2, 4, 9)
        prnt4excel(sw, "TXT", "Cri.Amt.Qtr.End", 2, 4, 10)
        prnt4excel(sw, "TXT", "NPA Reason", 2, 4, 11)
        prnt4excel(sw, "TXT", "AOD Due On", 2, 4, 12)
        prnt4excel(sw, "TXT", "Rep Schedule", 2, 5, 3)
        prnt4excel(sw, "TXT", "Rep Period", 2, 5, 4)
        prnt4excel(sw, "TXT", "Installment", 2, 5, 5)
        prnt4excel(sw, "TXT", "Rep Frequency", 2, 5, 6)
        prnt4excel(sw, "TXT", "First Inst Date", 2, 5, 8)
        prnt4excel(sw, "TXT", "Whether Rescheduled In Finacle", 2, 5, 9)
        prnt4excel(sw, "TXT", "Parties", 2, 6, 3)
        prnt4excel(sw, "TXT", "Cust ID", 2, 6, 4)
        prnt4excel(sw, "TXT", "Customer Name", 2, 6, 5)
        prnt4excel(sw, "TXT", "Relation", 2, 6, 8)
        prnt4excel(sw, "TXT", "Gold Loan", 2, 6, 10)
        prnt4excel(sw, "TXT", "Mobile No", 2, 6, 11)
        prnt4excel(sw, "TXT", "TotalDep", 2, 6, 12)
        prnt4excel(sw, "TXT", "Notice", 2, 7, 3)
        prnt4excel(sw, "TXT", "Notice Date", 2, 7, 4)
        prnt4excel(sw, "TXT", "Send To", 2, 7, 5)
        prnt4excel(sw, "TXT", "Notice Name", 2, 7, 8)
        prnt4excel(sw, "TXT", "NPA dt.", 2, 8, 1)
        prnt4excel(sw, "TXT", "A/c No", 2, 8, 2)
        prnt4excel(sw, "TXT", "Follow Up", 2, 8, 3)
        prnt4excel(sw, "TXT", "Date", 2, 8, 4)
        prnt4excel(sw, "TXT", "Contacted", 2, 8, 5)
        prnt4excel(sw, "TXT", "Cont Type", 2, 8, 6)
        prnt4excel(sw, "TXT", "Initiated By", 2, 8, 7)
        prnt4excel(sw, "TXT", "Done By", 2, 8, 8)
        prnt4excel(sw, "TXT", "Response", 2, 8, 9)
        prnt4excel(sw, "hal", 3, 2, 3, 1, 8, 12)
        prnt4excel(sw, "FNTSIZ", 9, 2, 3, 1, 8, 12)
        prnt4excel(sw, "fntbold", "", 2, 3, 1, 8, 12)
        prnt4excel(sw, "MERGE", "", 2, 5, 6, 5, 7)
        prnt4excel(sw, "MERGE", "", 2, 5, 9, 5, 12)
        prnt4excel(sw, "MERGE", "", 2, 6, 5, 6, 7)
        prnt4excel(sw, "MERGE", "", 2, 6, 8, 6, 9)
        prnt4excel(sw, "MERGE", "", 2, 7, 5, 7, 7)
        prnt4excel(sw, "MERGE", "", 2, 7, 8, 7, 12)
        prnt4excel(sw, "MERGE", "", 2, 8, 9, 8, 12)

        ROWNUM = 8

        '================= WRITING DATA TO EXCEL ================================
        totnum = 0
        totbal = 0
        totcritic = 0

        While drx.Read
            npamainnum = drx(24)
            Select Case npamainnum
                Case 1
                    acno = drx(0)
                    solid = drx(2)
                    TEXT11 = drx(20)
                    DATE11 = drx(6)
                    DATE12 = drx(7)
                    RO_DT = drx(21) & "/" & drx(22)
                    cugland = Trim(Val(solid) - 40000 + 4000)
                    cugmob = Trim(Val(solid) - 40000 + 9400999000)
                    TEXT14 = drx(23)
                    DATE13 = drx(8)
                    STFPOS = drx(13)
                    NPADT = drx(25)
                    ROWNUM = ROWNUM + 1

                    If ROWNUM <> 9 Then
                        prnt4excel(sw, "txt", PREVDT, 2, ROWNUM, 1)
                        prnt4excel(sw, "txt", PREVACNO, 2, ROWNUM, 2)
                        prnt4excel(sw, "nf", "################", 2, ROWNUM, 2, ROWNUM, 2)
                        prnt4excel(sw, "txt", "Remarks", 2, ROWNUM, 3, ROWNUM, 3)
                        prnt4excel(sw, "MERGE", "", 2, ROWNUM, 4, ROWNUM, 12)
                        ROWNUM = ROWNUM + 1
                    End If

                    processmessage(Mid(WORK_BOOK_NAME, 9) & ":All in one : Writing row " & ROWNUM)
                    Application.DoEvents()

                    prnt4excel(sw, "TXT", NPADT, 2, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 2, ROWNUM, 2)
                    prnt4excel(sw, "TXT", solid, 2, ROWNUM, 3)
                    prnt4excel(sw, "TXT", UCase(TEXT11), 2, ROWNUM, 4)
                    prnt4excel(sw, "TXT", DATE11, 2, ROWNUM, 5)
                    prnt4excel(sw, "TXT", DATE12, 2, ROWNUM, 6)
                    prnt4excel(sw, "TXT", RO_DT, 2, ROWNUM, 7)
                    prnt4excel(sw, "TXT", cugland, 2, ROWNUM, 8)
                    prnt4excel(sw, "TXT", cugmob, 2, ROWNUM, 9)
                    prnt4excel(sw, "TXT", TEXT14, 2, ROWNUM, 10)
                    prnt4excel(sw, "TXT", DATE13, 2, ROWNUM, 11)
                    prnt4excel(sw, "TXT", STFPOS, 2, ROWNUM, 12)

                    PREVDT = NPADT
                    PREVACNO = acno

                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 2, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 6, ROWNUM, 6)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 7, ROWNUM, 7)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 9, ROWNUM, 9)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 10, ROWNUM, 10)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 11, ROWNUM, 11)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 12, ROWNUM, 12)
                    prnt4excel(sw, "FNTbold", "", 2, ROWNUM, 1, ROWNUM, 12)

                    acno = drx(0)
                    SCHEMECODE = drx(15)
                    DATE1 = drx(3)
                    NUMBER1 = drx(9)
                    NUMBER2 = drx(10)
                    DATE2 = drx(4)
                    NUMBER3 = drx(11)
                    NUMBER4 = drx(12)
                    TEXT1 = drx(14)
                    DATE4 = drx(5)
                    CRIAMTQTREND = CRIAMTQTREND = Math.Round(drx(28), 0, MidpointRounding.AwayFromZero)

                    ROWNUM = ROWNUM + 1

                    totnum = totnum + 1
                    totbal = totbal + NUMBER2
                    totcritic = totcritic + NUMBER4

                    prnt4excel(sw, "TXT", NPADT, 2, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 2, ROWNUM, 2)
                    prnt4excel(sw, "TXT", SCHEMECODE, 2, ROWNUM, 3)
                    prnt4excel(sw, "TXT", DATE1, 2, ROWNUM, 4)
                    prnt4excel(sw, "TXT", NUMBER1, 2, ROWNUM, 5)
                    prnt4excel(sw, "TXT", NUMBER2, 2, ROWNUM, 6)
                    prnt4excel(sw, "TXT", DATE2, 2, ROWNUM, 7)
                    prnt4excel(sw, "TXT", NUMBER3, 2, ROWNUM, 8)
                    prnt4excel(sw, "TXT", NUMBER4, 2, ROWNUM, 9)

                    CRITROW = ROWNUM

                    prnt4excel(sw, "TXT", CRIAMTQTREND, 2, ROWNUM, 10)
                    prnt4excel(sw, "TXT", TEXT1, 2, ROWNUM, 11)
                    prnt4excel(sw, "TXT", DATE4, 2, ROWNUM, 12)

                    If NUMBER4 > 0 Then
                        prnt4excel(sw, "COMM", "Balance to Crit.amt ratio=" & Format(NUMBER2 / NUMBER4, "fixed"), 2, CRITROW, 9)
                    End If

                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 2, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 6, ROWNUM, 6)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 7, ROWNUM, 7)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 9, ROWNUM, 9)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 10, ROWNUM, 10)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 11, ROWNUM, 11)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 12, ROWNUM, 12)

                Case 2  'S
                    acno = drx(0)
                    RePeriod = drx(9)
                    Installment = drx(10)
                    RepFrequency = drx(14)
                    FirstInstDate = drx(3)
                    WhetherRescheduled = drx(15)

                    ROWNUM = ROWNUM + 1

                    prnt4excel(sw, "MERGE", "", 2, ROWNUM, 6, ROWNUM, 7)
                    prnt4excel(sw, "MERGE", "", 2, ROWNUM, 9, ROWNUM, 12)
                    prnt4excel(sw, "TXT", NPADT, 2, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 2, ROWNUM, 2)
                    prnt4excel(sw, "TXT", "Rep Schedule", 2, ROWNUM, 3)
                    prnt4excel(sw, "TXT", RePeriod, 2, ROWNUM, 4)
                    prnt4excel(sw, "TXT", Installment, 2, ROWNUM, 5)
                    prnt4excel(sw, "TXT", RepFrequency, 2, ROWNUM, 6)
                    prnt4excel(sw, "TXT", FirstInstDate, 2, ROWNUM, 8)
                    prnt4excel(sw, "TXT", WhetherRescheduled, 2, ROWNUM, 9)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 2, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 6, ROWNUM, 6)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 9, ROWNUM, 9)

                Case 3   'C
                    acno = drx(0)
                    CustID = drx(18)
                    CustomerName = drx(15)
                    Relation = drx(14)
                    GoldLoan = drx(10)
                    MobileNo = drx(16)
                    TotalDep = drx(9)

                    ROWNUM = ROWNUM + 1

                    prnt4excel(sw, "MERGE", "", 2, ROWNUM, 5, ROWNUM, 7)
                    prnt4excel(sw, "MERGE", "", 2, ROWNUM, 8, ROWNUM, 9)
                    prnt4excel(sw, "TXT", NPADT, 2, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 2, ROWNUM, 2)
                    prnt4excel(sw, "TXT", "Parties", 2, ROWNUM, 3)
                    prnt4excel(sw, "TXT", CustID, 2, ROWNUM, 4)
                    prnt4excel(sw, "TXT", CustomerName, 2, ROWNUM, 5)
                    prnt4excel(sw, "TXT", Relation, 2, ROWNUM, 8)
                    prnt4excel(sw, "TXT", GoldLoan, 2, ROWNUM, 10)
                    prnt4excel(sw, "TXT", MobileNo, 2, ROWNUM, 11)
                    prnt4excel(sw, "TXT", TotalDep, 2, ROWNUM, 12)

                    If Relation = "A/C HOLDER" Then
                        If NUMBER4 > 0 Then
                            If TotalDep >= NUMBER4 Then
                                prnt4excel(sw, "BG", 13434825, 2, CRITROW, 9, CRITROW, 9)
                                prnt4excel(sw, "FNTBOLD", "", 2, CRITROW, 9, CRITROW, 9)
                            End If
                        End If

                        If CRIAMTQTREND > 0 Then
                            If TotalDep >= CRIAMTQTREND Then
                                prnt4excel(sw, "BG", 13434825, 2, CRITROW, 10, CRITROW, 10)
                                prnt4excel(sw, "FNTBOLD", "", 2, CRITROW, 10, CRITROW, 10)
                            End If
                        End If
                    End If

                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 2, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 10, ROWNUM, 10)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 11, ROWNUM, 11)
                    prnt4excel(sw, "nf", "##########", 2, ROWNUM, 12, ROWNUM, 12)

                Case 4   'N
                    acno = drx(0)
                    NoticeDate = drx(3)
                    SendTo = drx(14)
                    NoticeName = drx(15)

                    ROWNUM = ROWNUM + 1

                    prnt4excel(sw, "MERGE", "", 2, ROWNUM, 5, ROWNUM, 7)
                    prnt4excel(sw, "MERGE", "", 2, CRITROW, 8, CRITROW, 12)
                    prnt4excel(sw, "TXT", NPADT, 2, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 2, ROWNUM, 2)
                    prnt4excel(sw, "TXT", "Notice", 2, ROWNUM, 3)
                    prnt4excel(sw, "TXT", NoticeDate, 2, ROWNUM, 4)
                    prnt4excel(sw, "TXT", SendTo, 2, ROWNUM, 5)
                    prnt4excel(sw, "TXT", NoticeName, 2, ROWNUM, 8)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 2, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 11, ROWNUM, 11)

                Case 5   'F
                    acno = drx(0)
                    FUDate = drx(3)
                    Contacted = drx(18)

                    If Contacted = "PARTY" Then Contacted = Contacted & "-" & drx(27)

                    ContType = drx(15)
                    InitiatedBy = drx(14)
                    DoneBy = drx(17)
                    Response = drx(19)

                    ROWNUM = ROWNUM + 1

                    prnt4excel(sw, "MERGE", "", 2, ROWNUM, 9, ROWNUM, 12)
                    prnt4excel(sw, "TXT", NPADT, 2, ROWNUM, 1)
                    prnt4excel(sw, "TXT", acno, 2, ROWNUM, 2)
                    prnt4excel(sw, "TXT", "FollowUp", 2, ROWNUM, 3)
                    prnt4excel(sw, "TXT", FUDate, 2, ROWNUM, 4)
                    prnt4excel(sw, "TXT", Contacted, 2, ROWNUM, 5)
                    prnt4excel(sw, "TXT", ContType, 2, ROWNUM, 6)
                    prnt4excel(sw, "TXT", InitiatedBy, 2, ROWNUM, 7)
                    prnt4excel(sw, "TXT", DoneBy, 2, ROWNUM, 8)
                    prnt4excel(sw, "TXT", Response, 2, ROWNUM, 9)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 1, ROWNUM, 1)
                    prnt4excel(sw, "nf", "################", 2, ROWNUM, 2, ROWNUM, 2)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 3, ROWNUM, 3)
                    prnt4excel(sw, "nf", "DD-MM-YYYY", 2, ROWNUM, 4, ROWNUM, 4)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 5, ROWNUM, 5)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 8, ROWNUM, 8)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 10, ROWNUM, 10)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 11, ROWNUM, 11)
                    prnt4excel(sw, "nf", "@", 2, ROWNUM, 12, ROWNUM, 12)
                Case Else
            End Select
        End While
        drx.Close()


        '================= SETTING FNT SIZE OF DATA ================================

        prnt4excel(sw, "FNTSIZ", 9, 2, 3, 1, ROWNUM, 12)
        prnt4excel(sw, "FNTNAM", "Arial", 2, 3, 1, ROWNUM, 12)
        prnt4excel(sw, "merge", "", 2, 3, 1, 7, 2)
        prnt4excel(sw, "hal", 2, 2, 3, 1, 3, 1)
        prnt4excel(sw, "val", 2, 2, 3, 1, 3, 1)
        prnt4excel(sw, "txt", "No of Accounts : " & totnum & Chr(10) & "Bal O/s : " & Format(totbal / 100000, "fixed") & " Lakhs" & Chr(10) & "Crit Amt : " & Format(totcritic / 100000, "fixed") & " Lakhs", 2, 3, 1, 3, 1)
        prnt4excel(sw, "fntbold", "", 2, 3, 1, 3, 1)
        prnt4excel(sw, "FNTSIZ", 11, 2, 3, 1, 3, 1)
        prnt4excel(sw, "hal", 2, 2, 9, 1, ROWNUM, 12)
        prnt4excel(sw, "FP", "", 2, 9, 1)

        '================= END OF SHEET 2 ==========================================================================================================


        '================= BEGINNING OF SHEET 3 ==========================================================================================================

        prnt4excel(sw, "shtname", "Account Master", 3)
        prnt4excel(sw, "colw", 11, 3, , 1)
        prnt4excel(sw, "colw", 15, 3, , 2)
        prnt4excel(sw, "colw", 11, 3, , 3)
        prnt4excel(sw, "colw", 11, 3, , 4)
        prnt4excel(sw, "colw", 9, 3, , 5)
        prnt4excel(sw, "colw", 9, 3, , 6)
        prnt4excel(sw, "colw", 11, 3, , 7)
        prnt4excel(sw, "colw", 10, 3, , 8)
        prnt4excel(sw, "colw", 10, 3, , 9)
        prnt4excel(sw, "colw", 10, 3, , 10)
        prnt4excel(sw, "colw", 21, 3, , 11)
        prnt4excel(sw, "colw", 11, 3, , 12)
        prnt4excel(sw, "TXT", "NPA dt.", 3, 1, 1)
        prnt4excel(sw, "TXT", "A/c No", 3, 1, 2)
        prnt4excel(sw, "TXT", "Scheme", 3, 1, 3)
        prnt4excel(sw, "TXT", "Open Date", 3, 1, 4)
        prnt4excel(sw, "TXT", "Loan", 3, 1, 5)
        prnt4excel(sw, "TXT", "Bal O/S", 3, 1, 6)
        prnt4excel(sw, "TXT", "Due Date", 3, 1, 7)
        prnt4excel(sw, "TXT", "Overdue", 3, 1, 8)
        prnt4excel(sw, "TXT", "Critical Amt", 3, 1, 9)
        prnt4excel(sw, "TXT", "Cri.Amt.Qtr.End", 3, 1, 10)
        prnt4excel(sw, "TXT", "NPA Reason", 3, 1, 11)
        prnt4excel(sw, "TXT", "AOD Due On", 3, 1, 12)
        prnt4excel(sw, "FNTSIZ", 9, 3, 1, 1, 1, 12)
        prnt4excel(sw, "FNTNAM", "Arial", 3, 1, 1, 1, 12)
        prnt4excel(sw, "fntbold", "", 3, 1, 1, 1, 12)
        prnt4excel(sw, "hal", 3, 3, 1, 1, 1, 12)

        Dim sql2 As String
        sql2 = "SELECT acno, TEXT2, DATE1, NUMBER1, NUMBER2,DATE2,nvl(NUMBER3,0), nvl(NUMBER4,0), TEXT1, DATE4, DATE3, nvl(NUMBER5,0) FROM C_MISADV WHERE NPAMAIN='M'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID =  " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO"
        Dim cmd2 As New OracleCommand(sql2, oracle_conn)
        Dim dr2 As OracleDataReader = cmd2.ExecuteReader()

        ROWNUM = 1
        While dr2.Read
            ROWNUM = ROWNUM + 1 ' Row no. of Excel

            '================= WRITING DATA TO EXCEL ================================

            processmessage(Mid(WORK_BOOK_NAME, 9) & ":Account Master : Writing row " & ROWNUM)
            Application.DoEvents()

            prnt4excel(sw, "TXT", dr2(10), 3, ROWNUM, 1)
            prnt4excel(sw, "TXT", dr2(0), 3, ROWNUM, 2)
            prnt4excel(sw, "TXT", dr2(1), 3, ROWNUM, 3)
            prnt4excel(sw, "TXT", dr2(2), 3, ROWNUM, 4)
            prnt4excel(sw, "TXT", dr2(3), 3, ROWNUM, 5)
            prnt4excel(sw, "TXT", dr2(4), 3, ROWNUM, 6)
            prnt4excel(sw, "TXT", dr2(5), 3, ROWNUM, 7)
            prnt4excel(sw, "TXT", dr2(6), 3, ROWNUM, 8)
            prnt4excel(sw, "TXT", dr2(7), 3, ROWNUM, 9)
            prnt4excel(sw, "TXT", dr2(11), 3, ROWNUM, 10)
            prnt4excel(sw, "TXT", dr2(8), 3, ROWNUM, 11)
            prnt4excel(sw, "TXT", dr2(9), 3, ROWNUM, 12)
        End While

        dr2.Close()

        prnt4excel(sw, "nf", "DD-MM-YYYY", 3, 2, 1, ROWNUM, 1)
        prnt4excel(sw, "nf", "################", 3, 2, 2, ROWNUM, 2)
        prnt4excel(sw, "nf", "@", 3, 2, 3, ROWNUM, 3)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 3, 2, 4, ROWNUM, 4)
        prnt4excel(sw, "nf", "##########", 3, 2, 5, ROWNUM, 5)
        prnt4excel(sw, "nf", "##########", 3, 2, 6, ROWNUM, 6)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 3, 2, 7, ROWNUM, 7)
        prnt4excel(sw, "nf", "##########", 3, 2, 8, ROWNUM, 8)
        prnt4excel(sw, "nf", "##########", 3, 2, 9, ROWNUM, 9)
        prnt4excel(sw, "nf", "##########", 3, 2, 10, ROWNUM, 10)
        prnt4excel(sw, "nf", "@", 3, 2, 11, ROWNUM, 11)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 3, 2, 12, ROWNUM, 12)
        prnt4excel(sw, "HAL", 3, 3, 2, 4, ROWNUM, 4)
        prnt4excel(sw, "HAL", 3, 3, 2, 7, ROWNUM, 7)
        prnt4excel(sw, "HAL", 3, 3, 2, 12, ROWNUM, 12)
        prnt4excel(sw, "HAL", 3, 3, 2, 1, ROWNUM, 1)

        '================= SETTING FNT SIZE OF DATA ================================

        prnt4excel(sw, "FNTSIZ", 9, 3, 2, 1, ROWNUM, 12)
        prnt4excel(sw, "FNTNAM", "Arial", 3, 2, 1, ROWNUM, 12)
        prnt4excel(sw, "FP", "", 3, 2, 1)

        '================= END OF SHEET 3 ==========================================================================================================


        '================= BEGINNING OF SHEET 4 ==========================================================================================================

        prnt4excel(sw, "shtname", "Repayment Schedule", 4)
        prnt4excel(sw, "TXT", "A/c No", 4, 1, 1)
        prnt4excel(sw, "TXT", "Rep Period", 4, 1, 2)
        prnt4excel(sw, "TXT", "Installment", 4, 1, 3)
        prnt4excel(sw, "TXT", "Rep Frequency", 4, 1, 4)
        prnt4excel(sw, "TXT", "First Inst Date", 4, 1, 5)
        prnt4excel(sw, "TXT", "Whether Rescheduled In Finacle", 4, 1, 6)
        prnt4excel(sw, "colw", 14, 4, , 1)
        prnt4excel(sw, "colw", 9, 4, , 2)
        prnt4excel(sw, "colw", 9, 4, , 3)
        prnt4excel(sw, "colw", 22, 4, , 4)
        prnt4excel(sw, "colw", 10, 4, , 5)
        prnt4excel(sw, "colw", 24, 4, , 6)
        prnt4excel(sw, "FNTSIZ", 9, 4, 1, 1, 1, 7)
        prnt4excel(sw, "FNTNAM", "Arial", 4, 1, 1, 1, 7)
        prnt4excel(sw, "fntbold", "", 4, 1, 1, 1, 7)
        prnt4excel(sw, "hal", 3, 4, 1, 1, 1, 7)

        Dim sql3 As String
        sql3 = "select acno,NUMBER1,NUMBER2,TEXT1,DATE1, TEXT2 from  C_MISADV WHERE NPAMAIN='S'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID =  " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO"
        Dim cmd3 As New OracleCommand(sql3, oracle_conn)
        Dim dr3 As OracleDataReader = cmd3.ExecuteReader()

        ROWNUM = 1
        While dr3.Read
            ROWNUM = ROWNUM + 1 ' Row no. of Excel

            '================= WRITING DATA TO EXCEL ================================
            processmessage(Mid(WORK_BOOK_NAME, 9) & ":Rpayment Schedule : Writing row " & ROWNUM)
            Application.DoEvents()

            prnt4excel(sw, "TXT", dr3(0), 4, ROWNUM, 1)
            prnt4excel(sw, "TXT", dr3(1), 4, ROWNUM, 2)
            prnt4excel(sw, "TXT", dr3(2), 4, ROWNUM, 3)
            prnt4excel(sw, "TXT", dr3(3), 4, ROWNUM, 4)
            prnt4excel(sw, "TXT", dr3(4), 4, ROWNUM, 5)
            prnt4excel(sw, "TXT", dr3(5), 4, ROWNUM, 6)
        End While

        dr3.Close()

        prnt4excel(sw, "nf", "################", 4, 2, 1, ROWNUM, 1)
        prnt4excel(sw, "nf", "##########", 4, 2, 2, ROWNUM, 2)
        prnt4excel(sw, "nf", "##########", 4, 2, 3, ROWNUM, 3)
        prnt4excel(sw, "nf", "@", 4, 2, 4, ROWNUM, 4)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 4, 2, 5, ROWNUM, 5)
        prnt4excel(sw, "nf", "@", 4, 2, 6, ROWNUM, 6)
        prnt4excel(sw, "HAL", 3, 4, 2, 5, ROWNUM, 5)

        '================= SETTING FNT SIZE OF DATA ================================

        prnt4excel(sw, "FNTSIZ", 9, 4, 2, 1, ROWNUM, 6)
        prnt4excel(sw, "FNTNAM", "Arial", 4, 2, 1, ROWNUM, 6)
        prnt4excel(sw, "FP", "", 4, 2, 1)

        '================= END OF SHEET 4 ==========================================================================================================


        '================= BEGINNING OF SHEET 5 ==========================================================================================================

        prnt4excel(sw, "shtname", "Parties", 5)
        prnt4excel(sw, "TXT", "A/c No", 5, 1, 1)
        prnt4excel(sw, "TXT", "Cust ID", 5, 1, 2)
        prnt4excel(sw, "TXT", "Customer Name", 5, 1, 3)
        prnt4excel(sw, "TXT", "Relation", 5, 1, 4)
        prnt4excel(sw, "TXT", "Gold Loan", 5, 1, 5)
        prnt4excel(sw, "TXT", "Mobile No", 5, 1, 6)
        prnt4excel(sw, "TXT", "Total Dep", 5, 1, 7)
        prnt4excel(sw, "colw", 14, 5, , 1)
        prnt4excel(sw, "colw", 10, 5, , 2)
        prnt4excel(sw, "colw", 32, 5, , 3)
        prnt4excel(sw, "colw", 15, 5, , 4)
        prnt4excel(sw, "colw", 9, 5, , 5)
        prnt4excel(sw, "colw", 12, 5, , 6)
        prnt4excel(sw, "colw", 10, 5, , 7)
        prnt4excel(sw, "FNTSIZ", 9, 5, 1, 1, 1, 7)
        prnt4excel(sw, "FNTNAM", "Arial", 5, 1, 1, 1, 7)
        prnt4excel(sw, "fntbold", "", 5, 1, 1, 1, 7)
        prnt4excel(sw, "hal", 3, 5, 1, 1, 1, 7)

        Dim sql4 As String
        sql4 = "select acno, TEXT5,	TEXT2,TEXT1,NUMBER2,nvl(TEXT3,' '),NUMBER1 from  C_MISADV WHERE NPAMAIN='C'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO, TEXT1"
        Dim cmd4 As New OracleCommand(sql4, oracle_conn)
        Dim dr4 As OracleDataReader = cmd4.ExecuteReader()

        ROWNUM = 1
        While dr4.Read

            ROWNUM = ROWNUM + 1 ' Row no. of Excel

            processmessage(Mid(WORK_BOOK_NAME, 9) & ":Parties : Writing row " & ROWNUM)
            Application.DoEvents()

            '================= WRITING DATA TO EXCEL ================================

            prnt4excel(sw, "TXT", dr4(0), 5, ROWNUM, 1)
            prnt4excel(sw, "TXT", dr4(1), 5, ROWNUM, 2)
            prnt4excel(sw, "TXT", dr4(2), 5, ROWNUM, 3)
            prnt4excel(sw, "TXT", dr4(3), 5, ROWNUM, 4)
            prnt4excel(sw, "TXT", dr4(4), 5, ROWNUM, 5)
            prnt4excel(sw, "TXT", dr4(5), 5, ROWNUM, 6)
            prnt4excel(sw, "TXT", dr4(6), 5, ROWNUM, 7)
        End While

        dr4.Close()

        prnt4excel(sw, "nf", "################", 5, 2, 1, ROWNUM, 1)
        prnt4excel(sw, "nf", "@", 5, 2, 2, ROWNUM, 2)
        prnt4excel(sw, "nf", "@", 5, 2, 3, ROWNUM, 3)
        prnt4excel(sw, "nf", "@", 5, 2, 4, ROWNUM, 4)
        prnt4excel(sw, "nf", "##########", 5, 2, 5, ROWNUM, 5)
        prnt4excel(sw, "nf", "@", 5, 2, 6, ROWNUM, 6)
        prnt4excel(sw, "nf", "##########", 5, 2, 7, ROWNUM, 7)

        '================= SETTING FNT SIZE OF DATA ================================

        prnt4excel(sw, "FNTSIZ", 9, 5, 2, 1, ROWNUM, 7)
        prnt4excel(sw, "FNTNAM", "Arial", 5, 2, 1, ROWNUM, 7)
        prnt4excel(sw, "FP", "", 5, 2, 1)

        '================= END OF SHEET 5 ==========================================================================================================


        '================= BEGINNING OF SHEET 6 ==========================================================================================================

        prnt4excel(sw, "shtname", "Notice", 6)
        prnt4excel(sw, "TXT", "A/c No", 6, 1, 1)
        prnt4excel(sw, "TXT", "Notice Date", 6, 1, 2)
        prnt4excel(sw, "TXT", "Send To", 6, 1, 3)
        prnt4excel(sw, "TXT", "Notice Name", 6, 1, 4)
        prnt4excel(sw, "colw", 14, 6, , 1)
        prnt4excel(sw, "colw", 10, 6, , 2)
        prnt4excel(sw, "colw", 21, 6, , 3)
        prnt4excel(sw, "colw", 32, 6, , 4)
        prnt4excel(sw, "FNTSIZ", 9, 6, 1, 1, 1, 5)
        prnt4excel(sw, "FNTNAM", "Arial", 6, 1, 1, 1, 5)
        prnt4excel(sw, "fntbold", "", 6, 1, 1, 1, 5)
        prnt4excel(sw, "hal", 3, 6, 1, 1, 1, 5)

        Dim sql5 As String
        sql5 = "select acno, nvl(DATE1, '01-JAN-1901'),NVL(TEXT1,' '),NVL(TEXT2,' ') from C_MISADV WHERE NPAMAIN='N'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO, DATE1"
        Dim cmd5 As New OracleCommand(sql5, oracle_conn)
        Dim dr5 As OracleDataReader = cmd5.ExecuteReader()

        ROWNUM = 1

        While dr5.Read
            ROWNUM = ROWNUM + 1 ' Row no. of Excel
            processmessage(Mid(WORK_BOOK_NAME, 9) & ":Notice : Writing row " & ROWNUM)
            Application.DoEvents()

            '================= WRITING DATA TO EXCEL ================================

            prnt4excel(sw, "TXT", dr5(0), 6, ROWNUM, 1)
            prnt4excel(sw, "TXT", dr5(1), 6, ROWNUM, 2)
            prnt4excel(sw, "TXT", dr5(2), 6, ROWNUM, 3)
            prnt4excel(sw, "TXT", dr5(3), 6, ROWNUM, 4)
        End While

        dr5.Close()

        prnt4excel(sw, "nf", "################", 6, 2, 1, ROWNUM, 1)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 6, 2, 2, ROWNUM, 2)
        prnt4excel(sw, "nf", "@", 6, 2, 3, ROWNUM, 3)
        prnt4excel(sw, "nf", "@", 6, 2, 4, ROWNUM, 4)
        prnt4excel(sw, "HAL", 3, 6, 2, 2, ROWNUM, 2)

        '================= SETTING FNT SIZE OF DATA ================================

        prnt4excel(sw, "FNTSIZ", 9, 6, 2, 1, ROWNUM, 4)
        prnt4excel(sw, "FNTNAM", "Arial", 6, 2, 1, ROWNUM, 4)
        prnt4excel(sw, "FP", "", 6, 2, 1)

        '================= END OF SHEET 6 ==========================================================================================================


        '================= BEGINNING OF SHEET 7 ==========================================================================================================

        prnt4excel(sw, "shtname", "Follow up", 7)
        prnt4excel(sw, "TXT", "A/c No", 7, 1, 1)
        prnt4excel(sw, "TXT", "Date", 7, 1, 2)
        prnt4excel(sw, "TXT", "Contacted", 7, 1, 3)
        prnt4excel(sw, "TXT", "Cont Type", 7, 1, 4)
        prnt4excel(sw, "TXT", "Initiated By", 7, 1, 5)
        prnt4excel(sw, "TXT", "Done By", 7, 1, 6)
        prnt4excel(sw, "TXT", "Response", 7, 1, 7)
        prnt4excel(sw, "colw", 14, 7, , 1)
        prnt4excel(sw, "colw", 9, 7, , 2)
        prnt4excel(sw, "colw", 22, 7, , 3)
        prnt4excel(sw, "colw", 15, 7, , 4)
        prnt4excel(sw, "colw", 14, 7, , 5)
        prnt4excel(sw, "colw", 20, 7, , 6)
        prnt4excel(sw, "colw", 48, 7, , 7)
        prnt4excel(sw, "FNTSIZ", 9, 7, 1, 1, 1, 7)
        prnt4excel(sw, "FNTNAM", "Arial", 7, 1, 1, 1, 7)
        prnt4excel(sw, "FNTbold", "", 7, 1, 1, 1, 7)
        prnt4excel(sw, "hal", 3, 7, 1, 1, 1, 7)

        Dim sql6 As String
        sql6 = "select NVL(acno,' '), DATE1, NVL(TEXT5,' '), NVL(TEXT6,' '), NVL(TEXT2,' '), NVL(TEXT1,' '), NVL(TEXT3,' '), NVL(TEXT4,' '), NVL(TEXT7,' ') from C_MISADV WHERE NPAMAIN='F'  AND ( NUMBER21 BETWEEN " & START_AMT & " AND " & END_AMT & ") AND ( SOLID IN ( SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY DATE20, NUMBER21 DESC, ACNO, DATE1"
        Dim cmd6 As New OracleCommand(sql6, oracle_conn)
        Dim dr6 As OracleDataReader = cmd6.ExecuteReader()

        ROWNUM = 1

        While dr6.Read
            ROWNUM = ROWNUM + 1 ' Row no. of Excel
            processmessage(Mid(WORK_BOOK_NAME, 9) & ":Followup : Writing row " & ROWNUM)
            Application.DoEvents()

            '================= WRITING DATA TO EXCEL ================================

            prnt4excel(sw, "TXT", dr6(0), 7, ROWNUM, 1)
            prnt4excel(sw, "TXT", dr6(1), 7, ROWNUM, 2)

            If dr6(2) = "PARTY" Then
                prnt4excel(sw, "TXT", dr6(2) & "-" & dr6(3), 7, ROWNUM, 3)
            Else
                prnt4excel(sw, "TXT", dr6(2), 7, ROWNUM, 3)
            End If

            prnt4excel(sw, "TXT", dr6(4), 7, ROWNUM, 4)
            prnt4excel(sw, "TXT", dr6(5), 7, ROWNUM, 5)
            prnt4excel(sw, "TXT", dr6(6) & "-" & dr6(7), 7, ROWNUM, 6)
            prnt4excel(sw, "TXT", dr6(8), 7, ROWNUM, 7)
        End While

        dr6.Close()

        prnt4excel(sw, "nf", "################", 7, 2, 1, ROWNUM, 1)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 7, 2, 2, ROWNUM, 2)
        prnt4excel(sw, "nf", "@", 7, 2, 3, ROWNUM, 3)
        prnt4excel(sw, "nf", "@", 7, 2, 4, ROWNUM, 4)
        prnt4excel(sw, "nf", "@", 7, 2, 5, ROWNUM, 5)
        prnt4excel(sw, "nf", "@", 7, 2, 6, ROWNUM, 6)
        prnt4excel(sw, "HAL", 3, 7, 2, 2, ROWNUM, 2)
        prnt4excel(sw, "FP", "", 7, 2, 1)

        '================= SETTING FNT SIZE OF DATA ================================

        prnt4excel(sw, "FNTSIZ", 9, 7, 2, 1, ROWNUM, 7)
        prnt4excel(sw, "FNTNAM", "Arial", 7, 2, 1, ROWNUM, 7)

        '================= END OF SHEET 7 ==========================================================================================================

        '================= BEGINNING OF SHEET 8 ==========================================================================================================

        prnt4excel(sw, "shtname", "Branch Details", 8)
        prnt4excel(sw, "TXT", "SOLID", 8, 1, 1)
        prnt4excel(sw, "TXT", "Branch Name", 8, 1, 2)
        prnt4excel(sw, "TXT", "RO", 8, 1, 3)
        prnt4excel(sw, "TXT", "Dist.", 8, 1, 4)
        prnt4excel(sw, "TXT", "OPEN DATE", 8, 1, 5)
        prnt4excel(sw, "TXT", "ONLINE DATE", 8, 1, 6)
        prnt4excel(sw, "TXT", "MANAGER", 8, 1, 7)
        prnt4excel(sw, "TXT", "STAFF ID", 8, 2, 7)
        prnt4excel(sw, "TXT", "Name", 8, 2, 8)
        prnt4excel(sw, "TXT", "JOINED DATE", 8, 2, 9)
        prnt4excel(sw, "TXT", "No of staff", 8, 1, 10)
        prnt4excel(sw, "TXT", "No. of a/cs", 8, 1, 11)
        prnt4excel(sw, "TXT", "Crit. Amt.", 8, 1, 12)
        prnt4excel(sw, "TXT", "Crit.Amt.Qtr.End", 8, 1, 13)
        prnt4excel(sw, "TXT", "Overdue", 8, 1, 14)
        prnt4excel(sw, "TXT", "Bal o/s", 8, 1, 15)
        prnt4excel(sw, "colw", 5, 8, , 1)
        prnt4excel(sw, "colw", 26, 8, , 2)
        prnt4excel(sw, "colw", 7, 8, , 3)
        prnt4excel(sw, "colw", 7, 8, , 4)
        prnt4excel(sw, "colw", 9, 8, , 5)
        prnt4excel(sw, "colw", 10, 8, , 6)
        prnt4excel(sw, "colw", 6, 8, , 7)
        prnt4excel(sw, "colw", 25, 8, , 8)
        prnt4excel(sw, "colw", 10, 8, , 9)
        prnt4excel(sw, "colw", 7, 8, , 10)
        prnt4excel(sw, "colw", 5, 8, , 11)
        prnt4excel(sw, "colw", 7, 8, , 12)
        prnt4excel(sw, "colw", 8, 8, , 13)
        prnt4excel(sw, "colw", 8, 8, , 14)
        prnt4excel(sw, "colw", 8, 8, , 15)
        prnt4excel(sw, "FNTSIZ", 9, 8, 1, 1, 2, 15)
        prnt4excel(sw, "FNTNAM", "Arial", 8, 1, 1, 2, 15)
        prnt4excel(sw, "FNTbold", "", 8, 1, 1, 2, 15)
        prnt4excel(sw, "hal", 3, 8, 1, 1, 2, 15)
        prnt4excel(sw, "merge", "", 8, 1, 7, 1, 9)
        prnt4excel(sw, "merge", "", 8, 1, 1, 2, 1)
        prnt4excel(sw, "merge", "", 8, 1, 2, 2, 2)
        prnt4excel(sw, "merge", "", 8, 1, 3, 2, 3)
        prnt4excel(sw, "merge", "", 8, 1, 4, 2, 4)
        prnt4excel(sw, "merge", "", 8, 1, 5, 2, 5)
        prnt4excel(sw, "merge", "", 8, 1, 6, 2, 6)
        prnt4excel(sw, "merge", "", 8, 1, 10, 2, 10)
        prnt4excel(sw, "merge", "", 8, 1, 11, 2, 11)
        prnt4excel(sw, "merge", "", 8, 1, 12, 2, 12)
        prnt4excel(sw, "merge", "", 8, 1, 13, 2, 13)
        prnt4excel(sw, "merge", "", 8, 1, 14, 2, 14)
        prnt4excel(sw, "merge", "", 8, 1, 15, 2, 15)
        prnt4excel(sw, "wrap", "", 8, 1, 10, 1, 10)
        prnt4excel(sw, "wrap", "", 8, 1, 11, 1, 11)
        prnt4excel(sw, "wrap", "", 8, 1, 12, 1, 12)
        prnt4excel(sw, "wrap", "", 8, 1, 13, 1, 13)
        prnt4excel(sw, "wrap", "", 8, 1, 14, 1, 14)
        prnt4excel(sw, "wrap", "", 8, 1, 15, 1, 15)
        prnt4excel(sw, "val", 3, 8, 1, 1, 1, 1)
        prnt4excel(sw, "val", 3, 8, 1, 2, 1, 2)
        prnt4excel(sw, "val", 3, 8, 1, 3, 1, 3)
        prnt4excel(sw, "val", 3, 8, 1, 4, 1, 4)
        prnt4excel(sw, "val", 3, 8, 1, 5, 1, 5)
        prnt4excel(sw, "val", 3, 8, 1, 6, 1, 6)
        prnt4excel(sw, "val", 3, 8, 1, 7, 1, 7)
        prnt4excel(sw, "val", 3, 8, 1, 8, 1, 8)
        prnt4excel(sw, "val", 3, 8, 1, 9, 1, 9)
        prnt4excel(sw, "val", 3, 8, 1, 10, 1, 10)
        prnt4excel(sw, "val", 3, 8, 1, 11, 1, 11)
        prnt4excel(sw, "val", 3, 8, 1, 12, 1, 12)
        prnt4excel(sw, "val", 3, 8, 1, 13, 1, 13)
        prnt4excel(sw, "val", 3, 8, 1, 14, 1, 14)
        prnt4excel(sw, "val", 3, 8, 1, 15, 1, 15)
        Dim SQLY As String = "update c_misadv A SET (NUMBER11,NUMBER12,NUMBER13,NUMBER14,NUMBER15) = (select COUNT(1), sum(number2), sum(number3), sum(number4), sum(number5) FROM C_MISADV B WHERE B.NPAMAIN='M' AND A.SOLID = B.SOLID AND B.NUMBER2 BETWEEN " & START_AMT & " AND " & END_AMT & " GROUP BY SOLID) WHERE A.NPAMAIN = 'Z' AND A.SOLID IN (SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")"

        oracle_execute_non_query("ten", username, username, SQLY)

        Dim SQLZ As String = "UPDATE C_MISADV A SET (TEXT4) = '.' WHERE  A.NPAMAIN='Z' AND A.TEXT4 IS NULL"

        oracle_execute_non_query("ten", username, username, SQLZ)

        Dim sql7 As String
        sql7 = "SELECT solid,TEXT1,TEXT2,TEXT3,DATE1,DATE2,NUMBER1,text4,date3,NVL(NUMBER8,0), NVL(NUMBER11,0), NVL(NUMBER12,0), NVL(NUMBER13,0), NVL(NUMBER14,0), NVL(NUMBER15,0) from C_MISADV WHERE (NPAMAIN='Z') AND (NUMBER1 IS NOT NULL) AND (SOLID IN (SELECT SOL_ID FROM SST WHERE SET_ID = " & SOLSET & ")) ORDER BY SOLID"
        Dim cmd7 As New OracleCommand(sql7, oracle_conn)
        Dim dr7 As OracleDataReader = cmd7.ExecuteReader()

        ROWNUM = 2

        While dr7.Read
            ROWNUM = ROWNUM + 1 ' Row no. of Excel
            processmessage(Mid(WORK_BOOK_NAME, 9) & ":Branch Details : Writing row " & ROWNUM)
            Application.DoEvents()

            '================= WRITING DATA TO EXCEL ================================

            prnt4excel(sw, "TXT", dr7(0), 8, ROWNUM, 1)
            prnt4excel(sw, "TXT", UCase(dr7(1)), 8, ROWNUM, 2)
            prnt4excel(sw, "TXT", dr7(2), 8, ROWNUM, 3)
            prnt4excel(sw, "TXT", dr7(3), 8, ROWNUM, 4)
            prnt4excel(sw, "TXT", dr7(4), 8, ROWNUM, 5)
            prnt4excel(sw, "TXT", dr7(5), 8, ROWNUM, 6)
            prnt4excel(sw, "TXT", dr7(6), 8, ROWNUM, 7)
            prnt4excel(sw, "TXT", dr7(7), 8, ROWNUM, 8)
            prnt4excel(sw, "TXT", dr7(8), 8, ROWNUM, 9)
            prnt4excel(sw, "TXT", dr7(9), 8, ROWNUM, 10)
            prnt4excel(sw, "TXT", dr7(10), 8, ROWNUM, 11)
            prnt4excel(sw, "TXT", dr7(13), 8, ROWNUM, 12)
            prnt4excel(sw, "TXT", dr7(14), 8, ROWNUM, 13)
            prnt4excel(sw, "TXT", dr7(12), 8, ROWNUM, 14)
            prnt4excel(sw, "TXT", dr7(11), 8, ROWNUM, 15)
        End While

        dr7.Close()

        prnt4excel(sw, "nf", "@", 8, 3, 1, ROWNUM, 1)
        prnt4excel(sw, "nf", "@", 8, 3, 2, ROWNUM, 2)
        prnt4excel(sw, "nf", "@", 8, 3, 3, ROWNUM, 3)
        prnt4excel(sw, "nf", "@", 8, 3, 4, ROWNUM, 4)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 8, 3, 5, ROWNUM, 5)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 8, 3, 6, ROWNUM, 6)
        prnt4excel(sw, "nf", "@", 8, 3, 7, ROWNUM, 7)
        prnt4excel(sw, "nf", "@", 8, 3, 8, ROWNUM, 8)
        prnt4excel(sw, "nf", "DD-MM-YYYY", 8, 3, 9, ROWNUM, 9)
        prnt4excel(sw, "nf", "###", 8, 3, 10, ROWNUM, 10)
        prnt4excel(sw, "nf", "###", 8, 3, 11, ROWNUM, 11)
        prnt4excel(sw, "nf", "###########", 8, 3, 12, ROWNUM, 12)
        prnt4excel(sw, "nf", "###########", 8, 3, 13, ROWNUM, 13)
        prnt4excel(sw, "nf", "###########", 8, 3, 14, ROWNUM, 14)
        prnt4excel(sw, "nf", "###########", 8, 3, 15, ROWNUM, 15)
        prnt4excel(sw, "HAL", 3, 8, 3, 5, ROWNUM, 5)
        prnt4excel(sw, "HAL", 3, 8, 3, 6, ROWNUM, 6)
        prnt4excel(sw, "HAL", 3, 8, 3, 9, ROWNUM, 9)

        '================= SETTING FNT SIZE OF DATA ================================

        prnt4excel(sw, "FNTSIZ", 9, 8, 3, 1, ROWNUM, 15)
        prnt4excel(sw, "FNTNAM", "Arial", 8, 3, 1, ROWNUM, 15)
        prnt4excel(sw, "FP", "", 8, 3, 1)

        '================= END OF SHEET 8 ==========================================================================================================

        '================= NAMING, SAVING & CLOSING WORKBOOK ================================

        PREVDT = ""
        PREVACNO = ""

        processmessage(Mid(WORK_BOOK_NAME, 9) & ":Over.")
        Application.DoEvents()
        oracle_conn.Close()

        sw.Close()

    End Sub
    Sub option59()      'NPA Reports
        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String = "aa"
        Dim path As String
        Dim uconfirm As String = "N"
        Dim dumpname As String = "CBS\Franklin\mydump"
        Dim process_executed As String
        Dim part As Integer = 0


        '-------------------------------------------------------------------------------------------------------------------
        ' Importing Dump
        '------------------------------------------------------------------------------------------------------------------
        path = InputBox("Enter Path", "Enter Value", "D")
        dumpname = InputBox("Enter Dump name without extention", "Enter Value", "mydump")
        process_executed = InputBox("Enter the Process to be Excecuted", "Enter Value", "ALL")

        'process_executed = InputBox("Enter the Process to be Excecuted", "Enter Value", "GENERATE EMAIL")
        ''processmessage("Deleting existing data")
        ''oracle_execute_non_query("ten", username, username, "DROP TABLE C_SPECIALINFO3_AAAA")

        ''processmessage("Rename Table")
        ''oracle_execute_non_query("ten", username, username, "RENAME C_SPECIALINFO3_ZZZZ TO C_SPECIALINFO3_AAAA")
        '------------------------------------------------------------------------------
        If process_executed.ToUpper() = "ALL" Then
            part = 0
        ElseIf process_executed.ToUpper() = "GENERATE REPORTS" Then
            part = 1
        ElseIf process_executed.ToUpper() = "GENERATE EMAIL" Then
            part = 2
        Else
            MsgBox("Enter ALL to execute whole process, Generate Reports to Generated and email reports, Generate Email to Generate Emails")
        End If

        If part = 0 Then
            If Directory.Exists(Disk & ":\CBS\script") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory(Disk & ":\CBS\script")
            End If
            processmessage("Creating truncate file")
            Dim sw1 As StreamWriter = New StreamWriter(Disk & ":\CBS\script\TRUNCATE.bat")
            sw1.WriteLine("@echo off")
            sw1.WriteLine("sqlplus cbs/cbs@ten @D:\CBS\SCRIPT\table_truncate /nolog")
            sw1.Close()

            Dim sw11 As StreamWriter = New StreamWriter(Disk & ":\CBS\script\table_truncate.sql")
            sw11.WriteLine("DROP TABLE C_SPECIALINFO3_AAAA;")
            sw11.WriteLine("RENAME C_SPECIALINFO3_ZZZZ TO C_SPECIALINFO3_AAAA;")
            sw11.WriteLine("TRUNCATE TABLE C_MISADV;")
            sw11.WriteLine("TRUNCATE TABLE C_MISDEP;")
            sw11.WriteLine("COMMIT;")
            sw11.Close()

            processmessage("Creating import file")
            Dim sw2 As StreamWriter = New StreamWriter(Disk & ":\CBS\script\import.bat")
            sw2.WriteLine("@echo off")
            sw2.WriteLine("imp " & username & "/" & username & "@ten file=" & path & ":\" & dumpname & ".dmp  full=yes")
            sw2.Close()

            processmessage("Creating index file")
            Dim sw3 As StreamWriter = New StreamWriter(Disk & ":\CBS\script\create_index.sql")
            sw3.WriteLine("ALTER INDEX C_SPECIALINFO_ZZZZ_IDX1 RENAME TO C_SPECIALINFO_AAAA_IDX1;")
            sw3.WriteLine("ALTER INDEX C_SPECIALINFO_ZZZZ_IDX2 RENAME TO C_SPECIALINFO_AAAA_IDX2;")
            sw3.WriteLine("ALTER INDEX C_SPECIALINFO_ZZZZ_IDX3 RENAME TO C_SPECIALINFO_AAAA_IDX3;")
            sw3.WriteLine("CREATE INDEX C_SPECIALINFO_ZZZZ_IDX1 ON C_SPECIALINFO3_ZZZZ (SI1_SOLID);")
            sw3.WriteLine("CREATE INDEX C_SPECIALINFO_ZZZZ_IDX2 ON C_SPECIALINFO3_ZZZZ (SI1_ACID);")
            sw3.WriteLine("CREATE INDEX C_SPECIALINFO_ZZZZ_IDX3 ON C_SPECIALINFO3_ZZZZ (SI1_NPAMAIN);")
            sw3.WriteLine("COMMIT;")
            sw3.Close()

            Dim sw4 As StreamWriter = New StreamWriter(Disk & ":\CBS\script\create_index.bat")
            sw4.WriteLine("@echo off")
            sw4.WriteLine("sqlplus cbs/cbs@ten @D:\CBS\SCRIPT\create_index.sql /nolog")
            sw4.Close()

            Process.Start(Disk & ":\CBS\script\TRUNCATE.bat")
            uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
            If uconfirm <> "Y" Then
                MsgBox("Exiting application")
                Exit Sub
            End If

            processmessage("Importing dump")
            Process.Start(Disk & ":\CBS\script\import.bat")

            uconfirm = InputBox("Enter Y and press OK once the process is over", "Confirm", "Y")
            If uconfirm <> "Y" Then
                MsgBox("Exiting application")
                processmessage("Creating Index")

                Exit Sub
            End If

            Process.Start("D:\CBS\SCRIPT\create_index.bat")

            If uconfirm <> "Y" Then
                MsgBox("Exiting application")
                processmessage("Creating Index")

                Exit Sub
            End If

            part = 1
        End If

        '-------------------------------------------------------------------------------------------------------------------
        ' GENERATING REPORTS
        '------------------------------------------------------------------------------------------------------------------


        'CALLING THE PACKGE

        If part = 1 Then

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package -NPA")      'KYC 

            sql = "PKGEMAIL117.GENERATE_NPAREPORTS"
            Dim cmd15 As New OracleCommand(sql, conn)
            cmd15.CommandType = CommandType.StoredProcedure
            cmd15.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
            cmd15.ExecuteNonQuery()

            '-------------------------------------------------------------------------------------------------------------------
            ' CREATING REPORTS
            '------------------------------------------------------------------------------------------------------------------

            processmessage("Checking folder path")

            If Directory.Exists("c:\du\ALL") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ALL")
            End If

            If Directory.Exists("c:\du\ROEKM") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROEKM")
            End If

            If Directory.Exists("c:\du\ROKKD") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROKKD")
            End If

            If Directory.Exists("c:\du\ROKNR") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROKNR")
            End If

            If Directory.Exists("c:\du\ROKPT") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROKPT")
            End If

            If Directory.Exists("c:\du\ROKSD") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROKSD")
            End If

            If Directory.Exists("c:\du\ROKTM") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROKTM")
            End If

            If Directory.Exists("c:\du\ROMPM") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROMPM")
            End If

            If Directory.Exists("c:\du\ROTLY") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROTLY")
            End If

            If Directory.Exists("c:\du\ROTSR") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROTSR")
            End If

            If Directory.Exists("c:\du\ROTVM") Then
                tempvar = "aa"
            Else
                Directory.CreateDirectory("c:\du\ROTVM")
            End If

            sql = "SELECT DISTINCT KGB_RO FROM Z_KGB WHERE KGB_RO<>'HO' UNION SELECT 'ALL' KGB_RO FROM DUAL"
            Dim cmd54 As New OracleCommand(sql, conn)
            Dim dr As OracleDataReader = cmd54.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("KGB_RO").ToString

                processmessage("FRESHINFLOW")
                ' File -1 FRESHINFLOW
                Dim file1 As String = "c:\du\" & tempvar & "\" & tempvar & "_FRESH_INFLOW.txt"
                If File.Exists(file1) Then

                    File.Delete(file1)

                End If
                Dim sw01 As StreamWriter = New StreamWriter(file1)
                sql = "SELECT MEMO1 FROM C_MISDEP WHERE TEXT1='FRESHINFLOW' AND TEXT2='" & tempvar & "'  ORDER BY TEXT1,TEXT2,NUMBER2"
                Dim cmd1 As New OracleCommand(sql, conn)
                Dim dr1 As OracleDataReader = cmd1.ExecuteReader()
                While dr1.Read()
                    tempvar1 = dr1.Item("MEMO1").ToString
                    If tempvar1 <> "" Then
                        sw01.WriteLine(tempvar1)
                    End If
                End While
                dr1.Close()
                sw01.Close()

                ' File -2 PNPA90DAYS
                processmessage("PNPA90DAYS")
                Dim file2 As String = "c:\du\" & tempvar & "\" & tempvar & "_PNPA_90DAYS.txt"
                If File.Exists(file2) Then

                    File.Delete(file2)

                End If
                Dim sw02 As StreamWriter = New StreamWriter(file2)
                sql = "SELECT MEMO1 FROM C_MISDEP WHERE TEXT1='PNPA90DAYS' AND TEXT2='" & tempvar & "'  ORDER BY TEXT1,TEXT2,NUMBER2"
                Dim cmd2 As New OracleCommand(sql, conn)
                Dim dr2 As OracleDataReader = cmd2.ExecuteReader()
                While dr2.Read()
                    tempvar2 = dr2.Item("MEMO1").ToString
                    If tempvar2 <> "" Then
                        sw02.WriteLine(tempvar2)
                    End If
                End While
                dr2.Close()
                sw02.Close()

                ' File -3 PNPAQUARTEREND
                processmessage("PNPAQUARTEREND")
                Dim file3 As String = "c:\du\" & tempvar & "\" & tempvar & "_PNPA_QUARTER_END.txt"
                If File.Exists(file3) Then

                    File.Delete(file3)

                End If
                Dim sw03 As StreamWriter = New StreamWriter(file3)
                Dim tempvar3 As String

                sql = "SELECT MEMO1 FROM C_MISDEP WHERE TEXT1='PNPAQUARTEREND' AND TEXT2='" & tempvar & "'  ORDER BY TEXT1,TEXT2,NUMBER2"
                Dim cmd3 As New OracleCommand(sql, conn)
                Dim dr3 As OracleDataReader = cmd3.ExecuteReader()
                While dr3.Read()
                    tempvar3 = dr3.Item("MEMO1").ToString
                    If tempvar3 <> "" Then
                        sw03.WriteLine(tempvar3)
                    End If
                End While
                dr3.Close()
                sw03.Close()

                ' File -4 NONLPDNPA
                processmessage("NONLPDNPA")
                Dim file4 As String = "c:\du\" & tempvar & "\" & tempvar & "_NON_LPD_NPA.txt"
                If File.Exists(file4) Then

                    File.Delete(file4)

                End If
                Dim sw04 As StreamWriter = New StreamWriter(file4)
                Dim tempvar4 As String
                sql = "SELECT MEMO1 FROM C_MISDEP WHERE TEXT1='NONLPDNPA' AND TEXT2='" & tempvar & "'  ORDER BY TEXT1,TEXT2,NUMBER2"
                Dim cmd4 As New OracleCommand(sql, conn)
                Dim dr4 As OracleDataReader = cmd4.ExecuteReader()
                While dr4.Read()
                    tempvar4 = dr4.Item("MEMO1").ToString
                    If tempvar4 <> "" Then
                        sw04.WriteLine(tempvar4)
                    End If
                End While
                dr4.Close()
                sw04.Close()

                ' File -5 LPDNPA
                processmessage("LPDNPA")
                Dim file5 As String = "c:\du\" & tempvar & "\" & tempvar & "_LPD_NPA.txt"
                If File.Exists(file5) Then

                    File.Delete(file5)

                End If
                Dim sw5 As StreamWriter = New StreamWriter(file5)
                Dim tempvar5 As String
                sql = "SELECT MEMO1 FROM C_MISDEP WHERE TEXT1='LPDNPA' AND TEXT2='" & tempvar & "'  ORDER BY TEXT1,TEXT2,NUMBER2"
                Dim cmd5 As New OracleCommand(sql, conn)
                Dim dr5 As OracleDataReader = cmd5.ExecuteReader()
                While dr5.Read()
                    tempvar5 = dr5.Item("MEMO1").ToString
                    If tempvar5 <> "" Then
                        sw5.WriteLine(tempvar5)
                    End If
                End While
                dr5.Close()
                sw5.Close()

                ' File -6 EXPIRED
                processmessage("EXPIRED_CREDIT_LIMIT")
                Dim file6 As String = "c:\du\" & tempvar & "\" & tempvar & "_EXPIRED_CREDIT_LIMIT.txt"
                If File.Exists(file6) Then

                    File.Delete(file6)

                End If
                Dim sw6 As StreamWriter = New StreamWriter(file6)
                Dim tempvar6 As String
                sql = "SELECT MEMO1 FROM C_MISDEP WHERE TEXT1='EXPIRED' AND TEXT2='" & tempvar & "'  ORDER BY TEXT1,TEXT2,NUMBER2"
                Dim cmd6 As New OracleCommand(sql, conn)
                Dim dr6 As OracleDataReader = cmd6.ExecuteReader()
                While dr6.Read()
                    tempvar6 = dr6.Item("MEMO1").ToString
                    If tempvar6 <> "" Then
                        sw6.WriteLine(tempvar6)
                    End If
                End While
                dr6.Close()
                sw6.Close()

                ' File -7 EXPIRED_2MNTH
                processmessage("EXPIRING_DURING_NEXT_2MONTHS")
                Dim file7 As String = "c:\du\" & tempvar & "\" & tempvar & "_EXPIRING_DURING_NEXT_2MONTHS.txt"
                If File.Exists(file7) Then

                    File.Delete(file7)

                End If
                Dim sw7 As StreamWriter = New StreamWriter(file7)
                Dim tempvar7 As String
                sql = "SELECT MEMO1 FROM C_MISDEP WHERE TEXT1='EXPIRED_2MNTH' AND TEXT2='" & tempvar & "'  ORDER BY TEXT1,TEXT2,NUMBER2"
                Dim cmd7 As New OracleCommand(sql, conn)
                Dim dr7 As OracleDataReader = cmd7.ExecuteReader()
                While dr7.Read()
                    tempvar7 = dr7.Item("MEMO1").ToString
                    If tempvar7 <> "" Then
                        sw7.WriteLine(tempvar7)
                    End If
                End While
                dr7.Close()
                sw7.Close()

                ' File -8 LOANS 5 LAKH AND ABOVE
                processmessage("LOANS 5 LAKH AND ABOVE")
                Dim file8 As String = "c:\du\" & tempvar & "\" & tempvar & "_LOANS_5LAKH_AND_ABOVE.txt"
                If File.Exists(file8) Then

                    File.Delete(file8)

                End If
                Dim sw8 As StreamWriter = New StreamWriter(file8)
                Dim tempvar8 As String
                sql = "SELECT MEMO1 FROM C_MISDEP WHERE TEXT1='ALL_5LKH' AND TEXT2='" & tempvar & "'  ORDER BY TEXT1,TEXT2,NUMBER2"
                Dim cmd8 As New OracleCommand(sql, conn)
                Dim dr8 As OracleDataReader = cmd8.ExecuteReader()
                While dr8.Read()
                    tempvar8 = dr8.Item("MEMO1").ToString
                    If tempvar8 <> "" Then
                        sw8.WriteLine(tempvar8)
                    End If
                End While
                dr8.Close()
                sw8.Close()

            End While
            dr.Close()

            processmessage("Compresing -ALL")
            Dim source As String
            Dim destination As String
            Dim directoryname As String
            directoryname = "ALL"
            source = "C:\du\ALL"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROEKM")
            directoryname = "ROEKM"
            source = "C:\du\ROEKM"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROKKD")
            directoryname = "ROKKD"
            source = "C:\du\ROKKD"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROKNR")
            directoryname = "ROKNR"
            source = "C:\du\ROKNR"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROKPT")
            directoryname = "ROKPT"
            source = "C:\du\ROKPT"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROKSD")
            directoryname = "ROKSD"
            source = "C:\du\ROKSD"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROKTM")
            directoryname = "ROKTM"
            source = "C:\du\ROKTM"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROMPM")
            directoryname = "ROMPM"
            source = "C:\du\ROMPM"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROTLY")
            directoryname = "ROTLY"
            source = "C:\du\ROTLY"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROTSR")
            directoryname = "ROTSR"
            source = "C:\du\ROTSR"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)

            processmessage("Compresing -ROTVM")
            directoryname = "ROTVM"
            source = "C:\du\ROTVM"
            destination = "C:\du"
            compress(destination, directoryname, source)
            Application.DoEvents()
            Thread.Sleep(1000)
            ''-----------------------------------------------------------
            ''Dim fi As FileInfo = "C:\du\ALL"
            ''Compress_gzip("C:\du\ALL", "Y")

            ''Dim startPath As String = "c:\example\start"
            ''Dim zipPath As String = "c:\example\result.zip"
            ''Dim extractPath As String = "c:\example\extract"

            ''ZipFile.CreateFromDirectory(startPath, zipPath)

            '-------------------------------------------------------------------------------------------------------------------
            ' GENERATING EMAILS
            '------------------------------------------------------------------------------------------------------------------
            conn.Close()
            part = 2
        End If

        If part = 2 Then

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()
            processmessage("Deleting existing data")

            oracle_execute_non_query("ten", username, username, "truncate table z_email")

            processmessage("Package - Data ID - 1175")       'NPA In Out

            sql = "PKGEMAIL117.DATAID_1175"
            Dim cmd6 As New OracleCommand(sql, conn)
            cmd6.CommandType = CommandType.StoredProcedure
            cmd6.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
            cmd6.ExecuteNonQuery()

            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE CUSTOMERID = 'NPAINOUT' ORDER BY SOLID"
            display_in_File(sql, "C:\du\SMS_" & RptDate.ToString("ddMMyyyy") & ".txt")
            Process.Start("C:\du\SMS_" & RptDate.ToString("ddMMyyyy") & ".txt")

            processmessage("Package - Data ID - 1194")       'NPA Fresh In Flow

            sql = "PKGEMAIL119.DATAID_1194"
            Dim cmd7 As New OracleCommand(sql, conn)
            cmd7.CommandType = CommandType.StoredProcedure
            cmd7.Parameters.Add("PREVIOUSWORKINGDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
            cmd7.ExecuteNonQuery()

            'sendemail("smgbmis@gmail.com", "ten", username, username)
            sendemail("nt@kgbmis.in", "ten", username, username)

        Dim dirs1 As String() = Directory.GetFiles("C:\DU")
        Dim dir1 As String
        For Each dir1 In dirs1

            If dir1 = ("C:\DU\ALL.rar") Then
                    sendemail_npa_files("chairmankeralagb@gmail.com;nkkrishnankutty46876@gmail.com;haridasanv@gmail.com;srnair32474@gmail.com;crmwing.kgb@gmail.com;smgbrl93@gmail.com", "kgbcreditwing@gmail.com;kgbhomis@gmail.com;kgbmis1@gmail.com;franklinkf@gmail.com;udayakumarcv@gmail.com;sureshsmgb1@gmail.com;kgbitw@gmail.com;", "C:\DU\ALL.rar", "NPA DATA AS ON " & txtdate.Text & " - [KGB]")
            ElseIf dir1 = ("C:\DU\ROKPT.rar") Then
                sendemail_npa_files("rokpt.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROKPT.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROKPT]")
                sendemail_npa_files("krisbanathur@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROKPT.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROKPT]")
            ElseIf dir1 = ("C:\DU\ROKNR.rar") Then
                sendemail_npa_files("roknr.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROKNR.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROKNR]")
            ElseIf dir1 = ("C:\DU\ROTLY.rar") Then
                sendemail_npa_files("rotly.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROTLY.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROTLY]")
            ElseIf dir1 = ("C:\DU\ROKKD.rar") Then
                sendemail_npa_files("rokzd.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROKKD.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROKKD]")
            ElseIf dir1 = ("C:\DU\ROMPM.rar") Then
                sendemail_npa_files("ropma.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROMPM.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROMPM]")
            ElseIf dir1 = ("C:\DU\ROTSR.rar") Then
                sendemail_npa_files("rotsr.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROTSR.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROTSR]")
            ElseIf dir1 = ("C:\DU\ROKSD.rar") Then
                sendemail_npa_files("roksd.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROKSD.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROKSD]")
            ElseIf dir1 = ("C:\DU\ROTVM.rar") Then
                sendemail_npa_files("rotvm.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROTVM.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROTVM]")
            ElseIf dir1 = ("C:\DU\ROEKM.rar") Then
                sendemail_npa_files("roekm.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROEKM.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROEKM]")
            ElseIf dir1 = ("C:\DU\ROKTM.rar") Then
                sendemail_npa_files("roktm.kgb@gmail.com", "kgbmis1@gmail.com", "C:\DU\ROKTM.rar", "NPA DATA AS ON " & txtdate.Text & " - [ROKTM]")
            End If

        Next
        conn.Close()
        End If

        MsgBox("Process Completed")
    End Sub
    Sub option62()      'MASS NEFT AGRICULTURE DEPT

        Dim path As String = ""
        Dim filename As String = ""
        Dim process_executed As String
        Dim mobilenumber As String = ""
        Dim solid As String

        process_executed = InputBox("Enter the Process to be Excecuted", "1 - Mass neft file creation 2- Status update and Response file Creation")

        If process_executed = "1" Then
            path = InputBox("Enter Path", "Enter Value", "C:\DU")
            filename = InputBox("Enter file name with extention", "Enter file name")
            solid = InputBox("Enter Branch Code", "Enter Value", "40348")
            mobilenumber = InputBox("Enter Mobile number", "Enter Value", "9446393275")
            If path <> "C:\DU" Then
                MsgBox("File should be placed in C:\DU folder and path should be  C:\DU ", MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If

            If filename.Contains(".") = False Then
                MsgBox("Enter File name with extension", MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If

            If solid.Length <> 5 Then
                MsgBox("Enter valid SOLID", MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If

            If solid.Substring(0, 2) <> "40" Then
                MsgBox("Enter valid SOLID starts with 40", MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If

            If mobilenumber.Length <> 10 Then
                MsgBox("Enter valid mobile number", MsgBoxStyle.Critical, "Error")
                Exit Sub
            End If

            Dim dir As String = String.Concat(path, "\", filename)
            If File.Exists(dir) = False Then
                processmessage("")
                MsgBox(String.Concat("files doesnot exists in the folder ", path), MsgBoxStyle.Critical, "Error")
            Else
                processmessage("Uploading file")
                uploadfiledata_without_trim(dir, username, "Y")

            End If

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - PKGMASSNEFT")

            sql = "PKGMASSNEFT.UPLOAD_AGR_DEP_NEFT_ENTRIES"
            Dim cmd1 As New OracleCommand(sql, conn)
            cmd1.CommandType = CommandType.StoredProcedure
            cmd1.Parameters.Add("", OracleDbType.Varchar2, 8, Nothing, ParameterDirection.Input).Value = solid
            cmd1.Parameters.Add("", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = filename
            cmd1.ExecuteNonQuery()

            processmessage("File generation")
            sql = "SELECT COUNT(1) FROM C_MISPRINT"
            Dim cmd2 As New OracleCommand(sql, conn)
            Dim dr1 As OracleDataReader = cmd2.ExecuteReader()
            tempcount = 0
            If dr1.Read = True Then
                tempcount = dr1(0)
            End If

            dr1.Close()
            If tempcount > 0 Then

                Dim sw11 As StreamWriter = New StreamWriter(String.Concat(path, "\Error_", solid, ".txt"))
                sql = String.Concat("SELECT SERIALNO||'|'||REPORTDATA AS LINEDATA FROM C_MISPRINT")

                Dim cmd14 As New OracleCommand(sql, conn)
                Dim dr11 As OracleDataReader = cmd14.ExecuteReader()
                While dr11.Read()
                    tempvar = dr11.Item("LINEDATA").ToString
                    If tempvar <> "" Then
                        sw11.WriteLine(tempvar)
                    End If
                End While
                dr11.Close()
                sw11.Close()
                Exit Sub
            End If

            Dim sw As StreamWriter = New StreamWriter(String.Concat(path, "\Upload_", solid, ".txt"))
            'sql = String.Concat("SELECT '", solid, "1013050114#'|| SUBSTR(TO_CHAR(SYSDATE,'DD-MM-YYYY'),0,10)||'#'||PKGSMGBCOMMON.TWODECIMALFORMAT(TRANSFER_AMT)||'#'||BEN_IFSC_CODE||'#'||BEN_ACNO||'#'||BEN_NAME||'#'||BEN_BRANCH_NAME||'##'||ROWNUM||'#'|| '", mobilenumber, "'||'##'||RECORD_ID2||'###' AS LINEDATA from  Z_AGR_DEP_NEFT_ENTRIES WHERE UPPER(FILE_NAME) = UPPER('" & path & "\" & filename.ToString.ToUpper & "') AND REMARKS = 'A' ")
            sql = String.Concat("SELECT '", solid, "1013050114#'|| SUBSTR(TO_CHAR(SYSDATE,'DD-MM-YYYY'),0,10)||'#'||TRIM(SUBSTR(PKGSMGBCOMMON.TWODECIMALFORMAT(TRANSFER_AMT),0,20))||'#'||TRIM(SUBSTR(BEN_IFSC_CODE,0,11))||'#'||TRIM(SUBSTR(BEN_ACNO,0,35))||'#'||TRIM(SUBSTR(BEN_NAME,0,100))||'#'||TRIM(SUBSTR(BEN_BRANCH_NAME,0,35))||'##'||ROWNUM||'#'|| '9446393275'||'##'||TRIM(SUBSTR(RECORD_ID2,0,33))||'###' AS LINEDATA from  Z_AGR_DEP_NEFT_ENTRIES WHERE UPPER(FILE_NAME) = UPPER('" & path & "\" & filename.ToString.ToUpper & "') AND REMARKS = 'A' ")

            Dim cmd4 As New OracleCommand(sql, conn)
            Dim dr As OracleDataReader = cmd4.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("LINEDATA").ToString
                If tempvar <> "" Then
                    sw.WriteLine(tempvar)
                End If
            End While
            dr.Close()
            sw.Close()

            Dim sw1 As StreamWriter = New StreamWriter(String.Concat(path, "\MASSUPLD_", solid, ".txt"))
            sql = String.Concat("SELECT  BEN_ACNO,TRANSFER_AMT,KGB_ACNAME,KGB_ACNO,RECORD_ID2 AS LINEDATA  from  Z_AGR_DEP_NEFT_ENTRIES WHERE UPPER(FILE_NAME) = UPPER('" & path & "\" & filename.ToUpper() & "') AND REMARKS = 'KGB' ")
            Dim cmd5 As New OracleCommand(sql, conn)
            Dim dr5 As OracleDataReader = cmd5.ExecuteReader()

            Dim LINEDATA As String = ""
            While dr5.Read()
                Dim credit_Acno As String = dr5(0).ToString
                Dim transferamnt As String = dr5(1).ToString
                Dim debit_acno As String = dr5(3).ToString
                Dim debitname As String = dr5(2).ToString
                Dim remarks_debit As String = dr5(4).ToString
                LINEDATA = ""
                LINEDATA = String.Concat(debit_acno, ",", transferamnt, ",D,", remarks_debit, ",,,,,")
                If LINEDATA <> "" Then
                    sw1.WriteLine(LINEDATA)
                End If

                LINEDATA = ""
                LINEDATA = String.Concat(credit_Acno, ",", transferamnt, ",C,", debit_acno, " ", debitname, ",,,,,")
                If LINEDATA <> "" Then
                    sw1.WriteLine(LINEDATA)
                End If

            End While
            dr5.Close()
            sw1.Close()

            MsgBox("File to send branch is generated in C:/du Foler.")

        ElseIf process_executed = 2 Then

            path = InputBox("Enter Path", "Enter Value", "C:\DU")
            filename = InputBox("Enter file name with extention", "Enter file name", "TABDATA_55444152.RPT")
            Dim neftdate As String
            neftdate = InputBox("Enter NEFT date in (dd-mm-yyyy) fromat", "Enter NEFT Date", "20-09-2016")
            If IsDate(neftdate) = False Then
                MsgBox("Enter valid date in dd-mm-yyyy")
            End If

            Dim upfilename As String
            upfilename = InputBox("Enter upload file name (File received from Department) with extention", "Enter file name")
            solid = InputBox("Enter Branch Code", "Enter Value", "40348")
            Dim dir As String = String.Concat(path, "\", filename)
            If File.Exists(dir) = False Then
                processmessage("")
                MsgBox(String.Concat("files doesnot exists in the folder ", path), MsgBoxStyle.Critical, "Error")
            Else
                processmessage("Uploading file")
                uploadfiledata_without_trim(dir, username, "Y")

            End If
            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            processmessage("Package - PKGMASSNEFT")

            sql = "PKGMASSNEFT.UPDATE_AGR_DEP_NEFT_ENTRIES"
            Dim cmd1 As New OracleCommand(sql, conn)
            cmd1.CommandType = CommandType.StoredProcedure
            cmd1.Parameters.Add("", OracleDbType.Varchar2, 8, Nothing, ParameterDirection.Input).Value = solid
            cmd1.Parameters.Add("", OracleDbType.Varchar2, 100, Nothing, ParameterDirection.Input).Value = upfilename
            cmd1.Parameters.Add("", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = neftdate
            cmd1.ExecuteNonQuery()


            processmessage("Response File generation")
            sql = "SELECT COUNT(1) FROM C_MISPRINT"
            Dim cmd2 As New OracleCommand(sql, conn)
            Dim dr1 As OracleDataReader = cmd2.ExecuteReader()
            tempcount = 0
            If dr1.Read = True Then
                tempcount = dr1(0)
            End If

            dr1.Close()
            If tempcount > 0 Then

                Dim sw11 As StreamWriter = New StreamWriter(String.Concat(path, "\Error_", solid, ".txt"))
                sql = String.Concat("SELECT SERIALNO||'|'||REPORTDATA AS LINEDATA FROM C_MISPRINT")

                Dim cmd14 As New OracleCommand(sql, conn)
                Dim dr11 As OracleDataReader = cmd14.ExecuteReader()
                While dr11.Read()
                    tempvar = dr11.Item("LINEDATA").ToString
                    If tempvar <> "" Then
                        sw11.WriteLine(tempvar)
                    End If
                End While
                dr11.Close()
                sw11.Close()
                Exit Sub
            End If

            Dim sw As StreamWriter = New StreamWriter(String.Concat(path, "\", upfilename.Replace(".txt", ""), "-Return (1).txt"))
            sql = String.Concat("SELECT RECORD_ID1||'|'||RECORD_ID2||'|'||KGB_ACNAME||'|'||KGB_ACNO||'|'||TRANSFER_AMT||'|'||REPLACE(TO_CHAR(FILE_RECORD_DATE,'DD-MM-YYYY'),'-','')||'|'||BEN_NAME||'|'||BEN_IFSC_CODE||'|'||BEN_BANK_NAME ||'|'||BEN_BRANCH_NAME||'|'||BEN_AC_TYPE||'|'||BEN_ACNO||'|'||REF_NO||'|'||UTR_NO||'|'||NEFT_STATUS||'|'||TO_CHAR('" & neftdate & "')||'|'||STATUS_REASON AS LINEDATA FROM Z_AGR_DEP_NEFT_ENTRIES WHERE UPPER(FILE_NAME) = UPPER('" & path & "\" & upfilename & "')")
            Dim cmd4 As New OracleCommand(sql, conn)
            Dim dr As OracleDataReader = cmd4.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("LINEDATA").ToString
                If tempvar <> "" Then
                    sw.WriteLine(tempvar)
                End If
            End While
            dr.Close()
            sw.Close()
            conn.Close()
            conn.Dispose()
            MsgBox("Response file to send the department is generated in C:/du Foler.")
        Else
            MsgBox("Invalid option")

        End If

    End Sub
    Sub option63()      'KIOSK FILE UPLOADING
        Dim path As String = ""
        Dim filename As String = ""

        path = InputBox("Enter Path", "Enter Value", "C:\DU")
        filename = InputBox("Enter file name with extention", "Enter file name")

        Dim dir As String = String.Concat(path, "\", filename)
        If File.Exists(dir) = False Then
            processmessage("")
            MsgBox(String.Concat("files doesnot exists in the folder ", path), MsgBoxStyle.Critical, "Error")
        Else
            processmessage("Uploading file")
            uploadfiledata_without_trim(dir, username, "Y")
        End If

        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Truncating Table")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_USER_STATUS_CODES")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_BC_MASTER")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_USER_VILL_MAPPING")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_VILLAGE_MASTER")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_STAFF_MASTER")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_SUBDIST_MASTER")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_DISTRICT_MASTER")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_STATE_MASTER")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_USER_MASTER")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_AC_ENR_STAT_CODES")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_AC_ENR_MODES")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_FIN_TRAN")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_FT_SER_CODES")

        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_NON_FIN_TRAN")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_NF_TRAN_SER_CODES")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_ACC_ENROLL")
        oracle_execute_non_query("ten", username, username, "truncate table Z_KIOSK_FINGER_PRINT_MASTER")

        processmessage("Executing Package Uploading data to tables")
        sql = "PKGEMAIL120.UPLOAD_KIOSK_ENTRIES"
        Dim cmd As New OracleCommand(sql, conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.ExecuteNonQuery()

        processmessage("Package - KIOSK MODULE - ENROLLMENTS - WEEKLY PROGRESS REPORT")      'KIOSK MODULE - ENROLLMENTS - WEEKLY PROGRESS REPORT
        sql = "PKGEMAIL121.DATAID_1211"
        Dim cmd1 As New OracleCommand(sql, conn)
        cmd1.CommandType = CommandType.StoredProcedure
        cmd1.Parameters.Add("GASON", OracleDbType.Date, 10, Nothing, ParameterDirection.Input).Value = Date.Now
        cmd1.ExecuteNonQuery()

        processmessage("Package - KIOSK MODULE - FINANCIAL TRANSACTIONS - WEEKLY PROGRESS REPORT")      'KIOSK MODULE - FINANCIAL TRANSACTIONS - WEEKLY PROGRESS REPORT
        sql = "PKGEMAIL121.DATAID_1212"
        Dim cmd2 As New OracleCommand(sql, conn)
        cmd2.CommandType = CommandType.StoredProcedure
        cmd2.Parameters.Add("GASON", OracleDbType.Date, 10, Nothing, ParameterDirection.Input).Value = Date.Now
        cmd2.ExecuteNonQuery()

        processmessage("Package - KIOSK MODULE - NON FINANCIAL TRANSACTIOS - WEEKLY PROGRESS REPORT")      'KIOSK MODULE - NON FINANCIAL TRANSACTIOS - WEEKLY PROGRESS REPORT
        sql = "PKGEMAIL122.DATAID_1221"
        Dim cmd3 As New OracleCommand(sql, conn)
        cmd3.CommandType = CommandType.StoredProcedure
        cmd3.Parameters.Add("GASON", OracleDbType.Date, 10, Nothing, ParameterDirection.Input).Value = Date.Now
        cmd3.ExecuteNonQuery()

        processmessage("Package - KIOSK MODULE - PEFORMANCE OF AKSHAYA/USB/KGB - WEEKLY PROGRESS REPORT")      'KIOSK MODULE - PEFORMANCE OF AKSHAYA/USB/KGB - WEEKLY PROGRESS REPORT
        sql = "PKGEMAIL122.DATAID_1222"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.Parameters.Add("GASON", OracleDbType.Date, 10, Nothing, ParameterDirection.Input).Value = Date.Now
        cmd4.ExecuteNonQuery()

        processmessage("Sending email")

        Dim sql1 As String
        Dim sql2 As String
        Dim tmplength As Integer = 0
        Dim substringno As Integer = 0
        Dim substringstart As Integer = 1
        Dim Bodydata As String = ""

        tempcount = 0
        sql = "SELECT MAIL_DATAID,MAIL_DATASUBID,MAIL_TO,NVL(MAIL_CC,'A')MAIL_CC,NVL(MAIL_BCC,'A')MAIL_BCC,MAIL_SUBJECT FROM Z_EMAIL ORDER BY MAIL_DATAID, MAIL_DATASUBID"
        Dim cmd21 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd21.ExecuteReader()
        While dr.Read()
            Bodydata = ""
            tmplength = 0
            substringno = 0
            substringstart = 1

            sql1 = "SELECT LENGTH(MAIL_BODY) TEMPLENGTH FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
            Dim cmd22 As New OracleCommand(sql1, conn)
            Dim dr1 As OracleDataReader = cmd22.ExecuteReader()
            While dr1.Read()
                tmplength = dr1.Item("TEMPLENGTH")
                Do Until tmplength = 0

                    If tmplength >= 3000 Then

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & ",3000) ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()

                    Else

                        substringno = tmplength

                        sql2 = "SELECT PKGSMGBCOMMON.GETCLOBDATA(MAIL_BODY," & substringstart & "," & substringno & ") ABCD FROM Z_EMAIL WHERE MAIL_DATAID = " & dr.Item("MAIL_DATAID") & " AND MAIL_DATASUBID = '" & dr.Item("MAIL_DATASUBID") & "'"
                        Dim cmd23 As New OracleCommand(sql2, conn)
                        Dim dr2 As OracleDataReader = cmd23.ExecuteReader()
                        While dr2.Read()
                            substringno = substringno + 3000
                            substringstart = substringstart + 3000
                            tmplength = tmplength - 3000
                            Bodydata = Bodydata & dr2.Item("ABCD")
                        End While
                        dr2.Close()
                        tmplength = 0

                    End If

                Loop

            End While
            dr1.Close()

            Dim oApp As Outlook._Application
            oApp = New Outlook.Application()
            Dim outlooksendfromaccount As String
            Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

            outlooksendfromaccount = "fag@kgbmis.in"

            Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)
            newMail.To = dr.Item("MAIL_TO")
            If dr.Item("MAIL_CC") <> "A" Then
                newMail.CC = dr.Item("MAIL_CC")
            End If
            If dr.Item("MAIL_BCC") <> "A" Then
                newMail.BCC = dr.Item("MAIL_BCC")
            End If
            newMail.Subject = dr.Item("MAIL_SUBJECT")
            newMail.HTMLBody = Bodydata
            newMail.SendUsingAccount = account
            newMail.Send()
            tempcount = tempcount + 1

            processmessage("Sending Mail No - " & tempcount)

        End While
        dr.Close()
        processmessage("")

        MsgBox("EMails Sent Successfully", MsgBoxStyle.Information, "Process Completed")
        conn.Close()

    End Sub

    Sub option64()          'Weekly Transaction Mail

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            ' Delete existing data, if any, from c_du table

            processmessage("Deleting existing data")

            oracle_execute_non_query("ten", username, username, "truncate table z_email")

            ' Calling packages

            processmessage("Package - POSEMAIL")

            sql = "PKGEMAIL124.POSEMAIL"
            Dim cmd1 As New OracleCommand(sql, conn)
            cmd1.CommandType = CommandType.StoredProcedure
            cmd1.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd1.ExecuteNonQuery()

            processmessage("Package - MB DATA")

            sql = "PKGEMAIL124.MBDATA"
            Dim cmd3 As New OracleCommand(sql, conn)
            cmd3.CommandType = CommandType.StoredProcedure
            cmd3.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd3.ExecuteNonQuery()

            processmessage("Package - ATM DATA")

            sql = "PKGEMAIL124.ATMDATA"
            Dim cmd4 As New OracleCommand(sql, conn)
            cmd4.CommandType = CommandType.StoredProcedure
            cmd4.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd4.ExecuteNonQuery()

            processmessage("Package - NEFT RTGS")

            sql = "PKGEMAIL124.NEFTRTGS"
            Dim cmd5 As New OracleCommand(sql, conn)
            cmd5.CommandType = CommandType.StoredProcedure
            cmd5.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd5.ExecuteNonQuery()

            processmessage("Package - TRANSACTIONS")

            sql = "PKGEMAIL124.TRANSACTIONS"
            Dim cmd6 As New OracleCommand(sql, conn)
            cmd6.CommandType = CommandType.StoredProcedure
            cmd6.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd6.ExecuteNonQuery()

            processmessage("Package - BILLDESK")

            sql = "PKGEMAIL124.BILLDESK"
            Dim cmd7 As New OracleCommand(sql, conn)
            cmd7.CommandType = CommandType.StoredProcedure
            cmd7.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd7.ExecuteNonQuery()

            processmessage("Package - Active Mobile Numbers used for Mobile Banking")

            sql = "PKGEMAIL125.ACTIVEMOBILEBANKING"
            Dim cmd8 As New OracleCommand(sql, conn)
            cmd8.CommandType = CommandType.StoredProcedure
            cmd8.Parameters.Add("GASON", OracleDbType.Date, 10, Nothing, ParameterDirection.Input).Value = Date.Now
            cmd8.ExecuteNonQuery()

            processmessage("Package - Active ATM Cards")

            sql = "PKGEMAIL125.ACTIVEATMCARDS"
            Dim cmd9 As New OracleCommand(sql, conn)
            cmd9.CommandType = CommandType.StoredProcedure
            cmd9.Parameters.Add("GASON", OracleDbType.Date, 10, Nothing, ParameterDirection.Input).Value = Date.Now
            cmd9.ExecuteNonQuery()

            processmessage("Package - ATM Hits")

            sql = "PKGEMAIL125.ATMHITS"
            Dim cmd10 As New OracleCommand(sql, conn)
            cmd10.CommandType = CommandType.StoredProcedure
            cmd10.Parameters.Add("GASON", OracleDbType.Date, 10, Nothing, ParameterDirection.Input).Value = Date.Now
            cmd10.ExecuteNonQuery()

            processmessage("Package - SingleWindow Opened Accounts")

            sql = "PKGEMAIL124.SWACCOUNT"
            Dim cmd11 As New OracleCommand(sql, conn)
            cmd11.CommandType = CommandType.StoredProcedure
            cmd11.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd11.ExecuteNonQuery()

            processmessage("Package - NPC issue Progress")

            sql = "PKGEMAIL125.NPCPROGRESS"
            Dim cmd12 As New OracleCommand(sql, conn)
            cmd12.CommandType = CommandType.StoredProcedure
            cmd12.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd12.ExecuteNonQuery()

            processmessage("Package - ATM issue Progress")

            sql = "PKGEMAIL125.ATMCARDISSUEPROGRESS"
            Dim cmd13 As New OracleCommand(sql, conn)
            cmd13.CommandType = CommandType.StoredProcedure
            cmd13.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
            cmd13.ExecuteNonQuery()

            sendemail("mis@kgbmis.in", "ten", username, username)

        End If
    End Sub
    Sub option65()      'STAFF Upload file creation for Dash Board

        ' Checking whether  files exists

        Dim tempvar As String
        Dim tempcount As String = 0

        processmessage("Checking files")

        file1 = "c:\du\STAFF_NAME.TXT"
        file2 = "c:\du\STAFF_BM.TXT"
        file3 = "c:\du\BUS.TXT"

        checkfile(file1, "Place the File naming as STAFF_NAME.TXT")
        checkfile(file2, "Place the File naming as STAFF_BM.TXT")
        checkfile(file3, "Place the File naming as BUS.TXT")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")
        uploadfiledata(file3, username, "N")

        ' Connecting to oracle data base

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        ' Calling packages

        processmessage("Package - PKGDASHBOARD.STAFF_UPLOAD_FILES")

        sql = "PKGDASHBOARD.STAFF_UPLOAD_FILES"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Creating STAFF_MASTER_REMOVE.txt")

        tempvar = ""
        Dim sw1 As StreamWriter = New StreamWriter("c:/du/STAFF_MASTER_REMOVE.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO BETWEEN 1 AND 4 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd12 As New OracleCommand(sql, conn)
        Dim dr1 As OracleDataReader = cmd12.ExecuteReader()
        While dr1.Read()
            tempvar = dr1.Item("REPORTDATA")
            sw1.WriteLine(tempvar)
        End While
        dr1.Close()
        sw1.Close()

        processmessage("Creating STAFF_MASTER_UPLOAD.txt")

        tempvar = ""
        Dim sw2 As StreamWriter = New StreamWriter("c:/du/STAFF_MASTER_UPLOAD.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO BETWEEN 5 AND 8 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd13 As New OracleCommand(sql, conn)
        Dim dr2 As OracleDataReader = cmd13.ExecuteReader()
        While dr2.Read()
            tempvar = dr2.Item("REPORTDATA")
            sw2.WriteLine(tempvar)
        End While
        dr2.Close()
        sw2.Close()

        processmessage("Creating SOL_IN_CHARGE.txt")

        tempvar = ""
        Dim sw3 As StreamWriter = New StreamWriter("c:/du/SOL_IN_CHARGE.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO BETWEEN 9 AND 12 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd14 As New OracleCommand(sql, conn)
        Dim dr3 As OracleDataReader = cmd14.ExecuteReader()
        While dr3.Read()
            tempvar = dr3.Item("REPORTDATA")
            sw3.WriteLine(tempvar)
        End While
        dr3.Close()
        sw3.Close()

        processmessage("Creating STAFF_POS.txt")

        tempvar = ""
        Dim sw4 As StreamWriter = New StreamWriter("c:/du/STAFF_POS.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO BETWEEN 13 AND 16 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd15 As New OracleCommand(sql, conn)
        Dim dr4 As OracleDataReader = cmd15.ExecuteReader()
        While dr4.Read()
            tempvar = dr4.Item("REPORTDATA")
            sw4.WriteLine(tempvar)
        End While
        dr4.Close()
        sw4.Close()

        processmessage("Creating SOL_TRAN_UPLOAD.txt")

        tempvar = ""
        Dim sw5 As StreamWriter = New StreamWriter("c:/du/SOL_TRAN_UPLOAD.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO BETWEEN 17 AND 20 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd16 As New OracleCommand(sql, conn)
        Dim dr5 As OracleDataReader = cmd16.ExecuteReader()
        While dr5.Read()
            tempvar = dr5.Item("REPORTDATA")
            sw5.WriteLine(tempvar)
        End While
        dr5.Close()
        sw5.Close()

        processmessage("Creating SOL_TRAN_REMOVE.txt")

        tempvar = ""
        Dim sw6 As StreamWriter = New StreamWriter("c:/du/SOL_TRAN_REMOVE.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO BETWEEN 21 AND 24 ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd17 As New OracleCommand(sql, conn)
        Dim dr6 As OracleDataReader = cmd17.ExecuteReader()
        While dr6.Read()
            tempvar = dr6.Item("REPORTDATA")
            sw6.WriteLine(tempvar)
        End While
        dr6.Close()
        sw6.Close()
        'processmessage("")

        MsgBox("Upload file created successfully", MsgBoxStyle.Information, "Process Completed")

        conn.Close()
        conn.Dispose()

    End Sub
    Sub option67()      'Mobile Banking sms creation

        ' Checking whether  files exists

        Dim tempvar As String
        Dim tempcount As String = 0

        processmessage("Checking files")

        file1 = "c:\du\SMS.TXT"
        checkfile(file1, "Place the File naming as SMS.TXT")
        
        uploadfiledata(file1, username, "Y")
        
        ' Connecting to oracle data base

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        ' Calling packages

        processmessage("Package - PKGCREATESMS.CREATE_MB_SMS")

        sql = "PKGCREATESMS.CREATE_MB_SMS"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.ExecuteNonQuery()

        processmessage("Creating MB_SMS.txt")

        tempvar = ""
        Dim sw1 As StreamWriter = New StreamWriter("c:/du/MB_SMS.txt")
        sql = "SELECT REPORTDATA FROM C_MISPRINT ORDER BY SERIALNO,SUBSERIALNO"
        Dim cmd12 As New OracleCommand(sql, conn)
        Dim dr1 As OracleDataReader = cmd12.ExecuteReader()
        While dr1.Read()
            tempvar = dr1.Item("REPORTDATA")
            sw1.WriteLine(tempvar)
        End While
        dr1.Close()
        sw1.Close()

        MsgBox("Upload file created successfully", MsgBoxStyle.Information, "Process Completed")

        conn.Close()
        conn.Dispose()

    End Sub

    Sub option68()      'Transacion file creation for dash board

        ' Checking whether  files exists

        Dim dirs As String() = Directory.GetFiles("c:\du")
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0

        totalfiles = dirs.Length

        If totalfiles = 0 Then

            processmessage("")

            MsgBox("No files exists in the folder c:/du", MsgBoxStyle.Critical, "Error")

        Else

            For Each dir In dirs

                tempcount = tempcount + 1

                If tempcount = 1 Then

                    uploadfiledata_without_trim(dir, username, "Y")

                Else

                    uploadfiledata_without_trim(dir, username, "N")

                End If

            Next

            Dim sql As String
            Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
            Dim conn As New OracleConnection(oradb)
            conn.Open()

            ' Calling packages

            processmessage("Package - PKGDASHBOARD.TRANSACTION_DASHBOARD_UPLOAD")

            sql = "PKGDASHBOARD.TRANSACTION_DASHBOARD_UPLOAD"
            Dim cmd4 As New OracleCommand(sql, conn)
            cmd4.CommandType = CommandType.StoredProcedure
            cmd4.ExecuteNonQuery()

            processmessage("Creating 41001.txt")

            tempvar = ""
            Dim sw1 As StreamWriter = New StreamWriter("c:/du/41001.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 1 ORDER BY SUBSERIALNO"
            Dim cmd12 As New OracleCommand(sql, conn)
            Dim dr1 As OracleDataReader = cmd12.ExecuteReader()
            While dr1.Read()
                tempvar = dr1.Item("REPORTDATA")
                sw1.WriteLine(tempvar)
            End While
            dr1.Close()
            sw1.Close()

            processmessage("Creating 41002.txt")

            tempvar = ""
            Dim sw2 As StreamWriter = New StreamWriter("c:/du/41002.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 2 ORDER BY SUBSERIALNO"
            Dim cmd13 As New OracleCommand(sql, conn)
            Dim dr2 As OracleDataReader = cmd13.ExecuteReader()
            While dr2.Read()
                tempvar = dr2.Item("REPORTDATA")
                sw2.WriteLine(tempvar)
            End While
            dr2.Close()
            sw2.Close()

            processmessage("Creating 41003.txt")

            tempvar = ""
            Dim sw3 As StreamWriter = New StreamWriter("c:/du/41003.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 3 ORDER BY SUBSERIALNO"
            Dim cmd14 As New OracleCommand(sql, conn)
            Dim dr3 As OracleDataReader = cmd14.ExecuteReader()
            While dr3.Read()
                tempvar = dr3.Item("REPORTDATA")
                sw3.WriteLine(tempvar)
            End While
            dr3.Close()
            sw3.Close()

            processmessage("Creating 41004.txt")

            tempvar = ""
            Dim sw4 As StreamWriter = New StreamWriter("c:/du/41004.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 4 ORDER BY SUBSERIALNO"
            Dim cmd15 As New OracleCommand(sql, conn)
            Dim dr4 As OracleDataReader = cmd15.ExecuteReader()
            While dr4.Read()
                tempvar = dr4.Item("REPORTDATA")
                sw4.WriteLine(tempvar)
            End While
            dr4.Close()
            sw4.Close()

            processmessage("Creating 41005.txt")

            tempvar = ""
            Dim sw5 As StreamWriter = New StreamWriter("c:/du/41005.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 5 ORDER BY SUBSERIALNO"
            Dim cmd16 As New OracleCommand(sql, conn)
            Dim dr5 As OracleDataReader = cmd16.ExecuteReader()
            While dr5.Read()
                tempvar = dr5.Item("REPORTDATA")
                sw5.WriteLine(tempvar)
            End While
            dr5.Close()
            sw5.Close()

            processmessage("Creating 41006.txt")

            tempvar = ""
            Dim sw6 As StreamWriter = New StreamWriter("c:/du/41006.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 5 ORDER BY SUBSERIALNO"
            Dim cmd17 As New OracleCommand(sql, conn)
            Dim dr6 As OracleDataReader = cmd17.ExecuteReader()
            While dr6.Read()
                tempvar = dr6.Item("REPORTDATA")
                sw6.WriteLine(tempvar)
            End While
            dr6.Close()
            sw6.Close()

            processmessage("Creating 41007.txt")

            tempvar = ""
            Dim sw7 As StreamWriter = New StreamWriter("c:/du/41007.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 7 ORDER BY SUBSERIALNO"
            Dim cmd18 As New OracleCommand(sql, conn)
            Dim dr7 As OracleDataReader = cmd18.ExecuteReader()
            While dr7.Read()
                tempvar = dr7.Item("REPORTDATA")
                sw7.WriteLine(tempvar)
            End While
            dr7.Close()
            sw7.Close()

            processmessage("Creating 41008.txt")

            tempvar = ""
            Dim sw8 As StreamWriter = New StreamWriter("c:/du/41008.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 8 ORDER BY SUBSERIALNO"
            Dim cmd19 As New OracleCommand(sql, conn)
            Dim dr8 As OracleDataReader = cmd18.ExecuteReader()
            While dr8.Read()
                tempvar = dr8.Item("REPORTDATA")
                sw8.WriteLine(tempvar)
            End While
            dr8.Close()
            sw8.Close()

            processmessage("Creating 41009.txt")

            tempvar = ""
            Dim sw9 As StreamWriter = New StreamWriter("c:/du/41009.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT WHERE SERIALNO = 9 ORDER BY SUBSERIALNO"
            Dim cmd20 As New OracleCommand(sql, conn)
            Dim dr9 As OracleDataReader = cmd20.ExecuteReader()
            While dr9.Read()
                tempvar = dr9.Item("REPORTDATA")
                sw9.WriteLine(tempvar)
            End While
            dr9.Close()
            sw9.Close()

            MsgBox("Upload file created successfully", MsgBoxStyle.Information, "Process Completed")

            conn.Close()
            conn.Dispose()
        End If

    End Sub
    Private Sub sendemail_npa_files(ByVal SEND_TO As String, ByVal SEND_CC As String, ByVal ATTACH1 As String, ByVal SUBJECT As String)

        'Generating EMail
        'Add "Microsoft Outlook 15.0 Object Library" in Project >> Reference >> Com
        'Add the following in declaration part
        'Imports System.Runtime.InteropServices
        'Imports Outlook = Microsoft.Office.Interop.Outlook

        Dim oApp As Outlook._Application
        oApp = New Outlook.Application()
        Dim outlooksendfromaccount As String
        Dim newMail As Outlook.MailItem = DirectCast(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        Dim dirs As String() = Directory.GetFiles("C:\DU")
        'Dim dir As String

        outlooksendfromaccount = "smgbmis@gmail.com"

        Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)

        newMail.To = SEND_TO
        newMail.CC = SEND_CC
        newMail.Subject = SUBJECT
        newMail.HTMLBody = "<html><body><head><style>.normalandleft{TEXT-ALIGN: left; FONT-FAMILY: arial, helvetica, sans-serif; FONT-SIZE: 9pt; FONT-WEIGHT: normal;}</style></head><p class=normalandleft>Dear Sir,</p><p class=normalandleft>Enclosed please find the following files containing NPA related data as on " & txtdate.Text & ":</p><p class=normalandleft>1. FRESH_INFLOW.txt - New accounts marked as NPA during the previous working day.<br>2. PNPA_QUARTER_END.txt - Account wise list of PNPA Accounts falling due during the current quarter.<br>3. PNPA_90DAYS.txt - Account wise list of PNPA Accounts falling due during the next 90 days.<br>4. NON_LPD_NPA.txt - Account wise list of Non LPD NPA Accounts.<br>5. LPD_NPA.txt - Account wise list of LPD NPA Accounts.<br>6. EXPIRED_CREDIT_LIMIT.txt - Account wise list of expired accounts during the previous working day.<br>7. EXPIRING_DURING_NEXT_2MONTHS.txt - Account wise list of accounts going to expire within two months.<br>8. LOANS_5LAKH_AND_ABOVE.txt - Account wise list of accounts having loan amount 5 Lakh and above.<br></p><p class=normalandleft>All files are prepared in pipedelimited format which can be directly copied into excel and split using the delimiter.</p><p class=normalandleft>Yours faithfully,</p><p class=normalandleft>MIS Team</p></body></html>"
        If ATTACH1 <> "0" Then
            newMail.Attachments.Add(ATTACH1)
        End If
        newMail.SendUsingAccount = account
        newMail.Send()


    End Sub

End Class