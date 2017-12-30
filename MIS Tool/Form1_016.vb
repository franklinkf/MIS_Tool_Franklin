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


Public Class Form1

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
    Dim menuitems_count = 63

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

        ElseIf rptoption = 25 Then

            option25()

        ElseIf rptoption = 26 Then

            option26()

        ElseIf rptoption = 27 Then

            option27()

        ElseIf rptoption = 28 Then

            option28()

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

        End If

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
        menulist(14, 1) = "CIBIL Upload File Creation"
        menulist(14, 2) = "A"
        menulist(15, 0) = "16"
        menulist(15, 1) = "EMail Daily Reports"
        menulist(15, 2) = "A"
        menulist(16, 0) = "17"
        menulist(16, 1) = "NPCI Linked Aadhaar - Upload file creation"
        menulist(16, 2) = "A"
        menulist(17, 0) = "18"
        menulist(17, 1) = "Day end Emails"
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

        checkfile(file1, "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder")
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

        processmessage("Package - Data ID - 1012")      'NPA In Out

        sql = "PKGEMAIL101.DATAID_1012"
        Dim cmd7 As New OracleCommand(sql, conn)
        cmd7.CommandType = CommandType.StoredProcedure
        cmd7.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd7.ExecuteNonQuery()

        processmessage("Package - Data ID - 1013")      'Loans Opened

        sql = "PKGEMAIL101.DATAID_1013"
        Dim cmd8 As New OracleCommand(sql, conn)
        cmd8.CommandType = CommandType.StoredProcedure
        cmd8.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd8.ExecuteNonQuery()

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

        processmessage("Package - Data ID - 1114")      'CGTMSE Accounts Not Linked With CGPAN

        sql = "PKGEMAIL111.DATAID_1114"
        Dim cmd66 As New OracleCommand(sql, conn)
        cmd66.CommandType = CommandType.StoredProcedure
        cmd66.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd66.ExecuteNonQuery()

        processmessage("Package - Data ID - 1131")      'Clientele Base

        sql = "PKGEMAIL113.DATAID_1131"
        Dim cmd67 As New OracleCommand(sql, conn)
        cmd67.CommandType = CommandType.StoredProcedure
        cmd67.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd67.ExecuteNonQuery()

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

    Sub option43()       'Figures At A Glance

        ' Checking whether email.txt file exists

        processmessage("Checking files")

        file1 = "c:\du\email.txt"

        checkfile(file1, "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder")
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

        processmessage("Package - Data ID - 1014")      'Figures At Glance

        sql = "PKGEMAIL101.DATAID_1014"
        Dim cmd9 As New OracleCommand(sql, conn)
        cmd9.CommandType = CommandType.StoredProcedure
        cmd9.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "ALL"
        cmd9.ExecuteNonQuery()

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

            If Val(dr.Item("MAIL_DATASUBID")) >= 40301 Then
                outlooksendfromaccount = "kgbmis1@gmail.com"
            Else
                outlooksendfromaccount = "smgbmis1@gmail.com"
            End If

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

        sendemail("smgbmis2@gmail.com", "ten", username, username)

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

    Sub option15()          'CIBIL Upload File Creation

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

            Dim sw As StreamWriter = New StreamWriter(Disk & ":/cibil/cibil_non_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=2 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd11 As New OracleCommand(sql, conn)
            Dim dr As OracleDataReader = cmd11.ExecuteReader()
            While dr.Read()
                tempvar = dr.Item("REPORTDATA")
                sw.WriteLine(tempvar)
            End While
            dr.Close()
            sw.Close()

            processmessage("Creating CIBIL_Individual.txt")

            tempvar = ""
            sw = New StreamWriter(Disk & ":/cibil/cibil_individual.txt")
            sql = "SELECT REPORTDATA FROM C_MISPRINT  WHERE SUBSERIALNO=1 ORDER BY SERIALNO,SUBSERIALNO"
            Dim cmd12 As New OracleCommand(sql, conn)
            dr = cmd12.ExecuteReader()
            While dr.Read()
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

            processmessage("Creating AnnexureA_Non_Individual.txt")

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

            processmessage("Creating AnnexureA_Individual.txt")

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

        newMail.To = "smgbmis@gmail.com"
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


    Sub option18()      'Day end EMails

        ' Checking whether 40998,40995,40994,KYC.TXT files exists

        processmessage("Checking files")

        file1 = "c:\du\40994.txt"
        file2 = "c:\du\40995.txt"
        file3 = "c:\du\40998.txt"
        file4 = "c:\du\KYC.txt"

        checkfile(file1, "Rename the file 40994_XX-XX-XXXX_AC1.TXT as 40994.TXT and place in c:/du folder")
        checkfile(file2, "Rename the file 40995_XX-XX-XXXX_AC1.TXT as 40995.TXT and place in c:/du folder")
        checkfile(file3, "Rename the file 40998AC1.TXT as 40998.TXT and place in c:/du folder")
        checkfile(file4, "Rename the upload error file KYC_XXXXXX.TXT as KYC.TXT and place in c:/du folder")

        uploadfiledata(file1, username, "Y")
        uploadfiledata(file2, username, "N")
        uploadfiledata(file3, username, "N")
        uploadfiledata(file4, username, "N")

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


        sendemail("smgbmis3@gmail.com", "ten", username, username)

    End Sub

    Sub option48()      'Branch Intimation Letter

        ' Checking whether files exists in C:/DU folder

        processmessage("Checking files")

        file1 = "c:\du\DATA.txt"

        checkfile(file1, "Rename the data file as DATA.TXT and place in c:/du folder")

        uploadfiledata(file1, username, "Y")

        ' Delete existing data, if any, from c_du table

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

        sendemail("kgbcreditwing@gmail.com", "ten", username, username)

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

        outlooksendfromaccount = "smgbmis4@gmail.com"

        Dim account As Outlook.Account = GetAccountForEmailAddress(oApp, outlooksendfromaccount)

        newMail.To = "chairmansmgb@gmail.com;Thangavelu.pp@keralagbank.com;thangavelu.ppt@gmail.com;KrishnanKutty.NK@smgbank.com;nkkrishnankutty46876@gmail.com;srnair32474@gmail.com;haridasanv@gmail.com"
        newMail.CC = "smgbrokzd1@smgbank.com;smgbrokpt1@smgbank.com;smgbropma1@smgbank.com;smgbrotsr1@smgbank.com;smgbaotvm1@smgbank.com;smgbhoit1@smgbank.com;nmgbrotly@gmail.com;nmgbaoekm@gmail.com;kgb660@keralagbank.com;nmgbksd@gmail.com;kgb395@keralagbank.com;smgbhopd1@smgbank.com;ditnmgb@gmail.com;nmgbknrao@gmail.com"
        newMail.BCC = "smgbchairman@smgbank.com;kgbhomis@gmail.com;franklin.kf@smgbank.com;franklinkf@gmail.com;udayakumarcv@gmail.com;sureshsmgb@gmail.com;smgbmis1@gmail.com;krisambali@gmail.com;Sebastian.KP@smgbank.com;FunctionalTeam.SMGB@smgbank.com;VinodK.Pattelath@smgbank.com"
        newMail.Subject = "Business Review - Excel and Mail Merge Word File with data as on " & txtdate.Text
        newMail.HTMLBody = "<html><body><head><style>.normalandleft{TEXT-ALIGN: left; FONT-FAMILY: arial, helvetica, sans-serif; FONT-SIZE: 9pt; FONT-WEIGHT: normal;}</style></head><p class=normalandleft>Dear Sir,</p><p class=normalandleft>Enclosed please find the following files containing the business figures of eSMGB brancheas as on " & txtdate.Text & ":</p><p class=normalandleft>1. Business Review.xlsx - To view the figures by providing the branch code/RO Code<br>2. Business Review.docx - To print the figures of branches in batch using the inbuilt mail merge facility.<br>3. Business Review Data.txt - Data source for the mail merge word file.  No specific use with that file<br></p><p class=normalandleft>To view/print the data, Download the attachment (compressed file), extract it and place the files in " & Disk & ":\Business Review Report</p><p class=normalandleft>In addition to this, the following facilites are available:</p><p class=normalandleft>1. Business review figures of every Fridays are emailed to branches/RO/HO on the next day<br>2. Business review figure of any day is available in Finacle MIS Server under Report ID - HMISRPT 210<br><p class=normalandleft>Yours faithfully,</p><p class=normalandleft>MIS Team</p></body></html>"
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
    Sub option25()          'Create RO and Branch Folders and convert CIB Files

        Dim foldercreationpath As String = "c:\du"
        Dim sourcefilepath As String = "C:\DU\CSV"
        Dim sourcefileextention As String = "csv"
        Dim dirs As String() = Directory.GetFiles(sourcefilepath, "*." & sourcefileextention)
        Dim folders As String()
        Dim folder As String
        Dim dir As String
        Dim totalfiles As Integer
        Dim tempcount As Integer = 0
        Dim filename As String
        Dim solid As String
        Dim subfolders As String()
        Dim subfolder As String
        Dim destinationpath As String
        'Creating folders and subfolders
        createdistrictbranchfolders(foldercreationpath)
        totalfiles = dirs.Length
        If totalfiles = 0 Then
            processmessage("")
            MsgBox("No files exists in the folder " & sourcefilepath, MsgBoxStyle.Critical, "Error")
        Else
            For Each dir In dirs
                destinationpath = ""
                tempcount = tempcount + 1
                filename = GetFileName(dir)
                solid = filename.Substring(0, 5)
                folders = Directory.GetDirectories(foldercreationpath)
                For Each folder In folders
                    subfolders = Directory.GetDirectories(folder)
                    For Each subfolder In subfolders
                        If InStr(subfolder, solid) > 0 Then
                            destinationpath = subfolder
                        End If
                    Next
                Next
                CreateExcelFromCsvFile(sourcefilepath, filename, sourcefileextention)
                processmessage("Converting File No - " & tempcount)
                If destinationpath <> "" Then
                    My.Computer.FileSystem.CopyFile(sourcefilepath & "\" & filename & ".xls", destinationpath & "\" & filename & ".xls", Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
                    processmessage("Moving File No - " & tempcount)
                End If
            Next
            processmessage("")
            MsgBox("Conversion completed successfully", MsgBoxStyle.Information, "Process Completed")
        End If
    End Sub

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

    Public Sub CreateExcelFromCsvFile(ByVal strFolderPath As String, ByVal strFileName As String)
        Dim newFileName As String = "NewExcelFile.xls"
        Dim oExcelFile As Object
        ' Open Excel application object
        Try
            oExcelFile = GetObject(, "Excel.Application")
        Catch
            oExcelFile = CreateObject("Excel.Application")
        End Try
        oExcelFile.Visible = False
        oExcelFile.Workbooks.Open(strFolderPath + "\" + strFileName)
        ' Turn off message box so that we do not get any messages
        oExcelFile.DisplayAlerts = False
        ' Save the file as XLS file
        oExcelFile.ActiveWorkbook.SaveAs(Filename:=strFolderPath + "\" + newFileName, FileFormat:=Excel.XlFileFormat.xlExcel5, CreateBackup:=False)
        ' Close the workbook
        oExcelFile.ActiveWorkbook.Close(SaveChanges:=False)
        ' Turn the messages back on
        oExcelFile.DisplayAlerts = True
        ' Quit from Excel
        oExcelFile.Quit()
        ' Kill the variable
        oExcelFile = Nothing
    End Sub

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
    Public Sub CreateExcelFromCsvFile(ByVal strFolderPath As String, ByVal strFileName As String, ByVal strfileextension As String)
        Dim newFileName As String = strFileName & ".xls"
        Dim oExcelFile As Object
        Try
            oExcelFile = GetObject(, "Excel.Application")
        Catch
            oExcelFile = CreateObject("Excel.Application")
        End Try
        oExcelFile.Visible = False
        oExcelFile.Workbooks.Open(strFolderPath + "\" + strFileName + "." + strfileextension)
        oExcelFile.DisplayAlerts = False
        oExcelFile.ActiveWorkbook.SaveAs(Filename:=strFolderPath + "\" + newFileName, FileFormat:=Excel.XlFileFormat.xlExcel5, CreateBackup:=False)
        oExcelFile.ActiveWorkbook.Close(SaveChanges:=False)
        oExcelFile.DisplayAlerts = True
        oExcelFile.Quit()
        oExcelFile = Nothing
    End Sub
    Public Function GetFileName(ByVal filepath As String) As String
        'This Function Gets the name of a file without the path or extension.
        Dim slashindex As Integer = filepath.LastIndexOf("\")
        Dim dotindex As Integer = filepath.LastIndexOf(".")
        GetFileName = filepath.Substring(slashindex + 1, dotindex - slashindex - 1)
    End Function
    Private Sub formatexcel(ByVal filename)
        Dim oExel As Excel.Application
        Dim oWorkbook As Excel.Workbook
        Dim oWorksheet As Excel.Worksheet
        Dim oRange As Excel.Range
        Dim rCnt As Integer
        Dim cCnt As Integer
        Dim Obj As Object
        Dim sReplace As String = "ABC"
        oExel = CreateObject("Excel.Application")
        oWorkbook = oExel.Application.Workbooks.Open(filename)
        oExel.Application.Interactive = True
        oExel.Application.UserControl = True
        For Each oWorksheet In oExel.ActiveWorkbook.Worksheets
            oRange = oWorksheet.UsedRange
            For rCnt = 1 To oRange.Rows.Count
                For cCnt = 1 To oRange.Columns.Count
                    Obj = CType(oRange.Cells(rCnt, cCnt), Excel.Range).Text
                    If Obj <> Nothing Then
                        ' find and replace
                        'MessageBox.Show(Obj)
                    End If

                Next
            Next
        Next
        oWorkbook.Save()
        oWorkbook.Close()
        oExel.Quit()
        oExel = Nothing
    End Sub

    Sub option28()          'Generate Word file

        Dim solid As String
        Dim solname As String

        Dim lpdsuit As Integer
        Dim lpdrr As Integer
        Dim lpdothers As Integer

        Dim tot_npa As Integer
        Dim npa_without_action As Integer
        Dim npa_without_action_march As Integer

        Dim suit_pend As Integer
        Dim rr_pend As Integer
        Dim legal_entered As Integer


        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table, oTable1 As Word.Table
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph
        Dim oPara3 As Word.Paragraph, oPara4 As Word.Paragraph, oPara5 As Word.Paragraph, oPara6 As Word.Paragraph
        Dim oRng As Word.Range
        Dim count As Integer

        count = 0


        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Setting page margin
        oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.0)
        oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.0)
        oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.75)
        oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.75)

        'Justify
        oDoc.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify




        oDoc.Range.Font.Name = "Abadi MT Condensed Light"
        oDoc.Range.Font.Size = 5
        oDoc.Paragraphs.Style = "No Spacing"

        'Add a picture at the header
        'oDoc.Content.Application.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader
        'oDoc.Content.Application.Selection.Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\1.jpg")


        'Dim PIctureLocation As String = "E:\VBProject\1.jpg"  --->Defining picture location
        'oDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\1.jpg")


        'Add picture in footer
        ''oDoc.Content.Application.Selection.Fields.Add(Range:=oDoc.Content.Application.Selection.Range, Type:=CInt(Word.WdFieldType.wdFieldEmpty), Text:="page")
        'oDoc.Content.Application.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument
        'oDoc.Content.Application.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter
        'oDoc.Content.Application.Selection.Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\2.jpg")
        ''oDoc.Content.Application.Selection.TypeText(Text:="Martens")

        'return to the main document        
        ' oDoc.Content.Application.ActiveWindow.ActivePane.View.SeekView = CInt(Word.WdSeekView.wdSeekMainDocument)

        sql = "SELECT TEXT1 ,TEXT20,NUMBER1,NUMBER2,NUMBER3,NUMBER4,NUMBER5,NUMBER6,NUMBER7,NUMBER8,NUMBER9 FROM C_MISADV WHERE C_ACID = 'SOLNAME' ORDER BY TEXT1"
        Dim cmd4 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd4.ExecuteReader()


        While dr.Read()
            solid = dr("text1").ToString()
            solname = dr("text20").ToString()

            lpdsuit = dr("number1")
            lpdrr = dr("number2")
            lpdothers = dr("number3")

            tot_npa = dr("number4")
            npa_without_action = dr("number5")
            npa_without_action_march = dr("number6")

            suit_pend = dr("number7")
            rr_pend = dr("number8")
            legal_entered = dr("number9")

            'Inserting a picture file
            oDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\1.jpg")

            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Format.SpaceAfter = 5
            oPara1.Range.InsertParagraphAfter()

            oPara1 = oDoc.Content.Paragraphs.Add()
            oPara1.Range.Text = "The Branch Manager"
            oPara1.Format.SpaceAfter = 1
            oPara1.Style = "No Spacing"
            oPara1.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
            oPara1.Range.Font.Size = 10
            oPara1.Range.InsertParagraphAfter()
            oPara1.Range.Text = "Kerala Gramin Bank"
            oPara1.Format.SpaceAfter = 1
            oPara1.Range.InsertParagraphAfter()
            oPara1.Range.Text = solname
            oPara1.Format.SpaceAfter = 5  'Setting space.
            oPara1.Range.InsertParagraphAfter()

            'Insert a paragraph at the end of the document.
            '** \endofdoc is a predefined bookmark.
            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            'oPara2.Format.SpaceAfter = 25
            oPara2.Range.Text = "Sir,"
            oPara2.Style = "No Spacing"
            oPara2.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
            oPara2.Range.Font.Size = 10

            oPara2.Format.SpaceAfter = 5
            oPara2.Range.InsertParagraphAfter()

            oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara3.Range.Text = "Sub: NPA Accounts with no action"
            oPara3.Style = "No Spacing"
            oPara3.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
            oPara3.Range.Font.Size = 10
            oPara3.Format.SpaceAfter = 5
            oPara3.Range.InsertParagraphAfter()
            oPara3.Style = "No Spacing"
            oPara3.Range.Text = "Furnished here below are the action initiated accounts (LPD Suit, LPD RR, LPD others), total number of NPA accounts and the number of accounts marked as NPA before 01/04/2013 and lying without any recovery action."

            oPara3.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
            oPara3.Range.Font.Size = 10
            oPara3.Format.SpaceAfter = 1
            oPara3.Range.InsertParagraphAfter()

            oPara3 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara3.SpaceAfter = 2

            'Create a table with 8 rows and 2 columns
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 2

            For r = 1 To 8
                For c = 1 To 2

                    If r = 1 Then
                        oTable.Cell(r, c).Range.Font.Bold = True


                        If c = 1 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                            oTable.Cell(r, c).Range.Text = "Head"
                        End If

                        If c = 2 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = "No. of A\c"
                        End If

                    ElseIf r = 2 Then

                        If c = 1 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable.Cell(r, c).Range.Text = "LPD Suit"
                        Else
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable.Cell(r, c).Range.Text = lpdsuit
                        End If

                    ElseIf r = 3 Then

                        If c = 1 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = "LPD RR"
                        Else
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = lpdrr
                        End If


                    ElseIf r = 4 Then

                        If c = 1 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = "LPD Others"

                        Else
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = lpdothers
                        End If



                    ElseIf r = 5 Then
                        If c = 1 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = "Total LPD"
                        Else
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = lpdothers + lpdrr + lpdsuit
                        End If



                    ElseIf r = 6 Then
                        If c = 1 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = "Total NPA"
                        Else
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = tot_npa
                        End If


                    ElseIf r = 7 Then

                        If c = 1 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = "Of which, NPA Accounts lying without action"
                        Else
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = npa_without_action
                        End If


                    ElseIf r = 8 Then
                        If c = 1 Then
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable.Cell(r, c).Range.Text = "                 NPA Accounts marked before March 2013 lying without action"
                        Else
                            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter

                            oTable.Cell(r, c).Range.Text = npa_without_action_march
                        End If
                    Else
                        oTable.Cell(r, c).Range.Text = "r" & r & "c" & c
                    End If
                    oTable.Cell(r, c).Borders.Enable = True
                Next
            Next
            oTable.Columns.Item(1).Width = oWord.InchesToPoints(5.7)   'Change width of columns 1 & 2
            oTable.Columns.Item(2).Width = oWord.InchesToPoints(1.1)

            oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara4.Format.SpaceAfter = 2

            oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara4.Format.SpaceAfter = 5

            oPara4.Range.Text = "Data entry status of LPD accounts under LPD module as instructed vide circular 3/R&L/2013 dated 17/07/2013 is given below:"
            oPara4.Style = "No Spacing"
            oPara4.Range.Font.Name = "Calibri (Body)"
            oPara4.Range.Font.Size = 10

            '--To get bold and underline
            'oPara4.Range.Font.Bold = True
            'oPara4.Range.Font.Underline = True

            oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara4.Format.SpaceAfter = 2
            oPara4.Range.InsertParagraphAfter()

            oPara4.Range.Font.Bold = False
            oPara4.Range.Font.Underline = False

            oTable1 = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 4)
            oTable1.Range.ParagraphFormat.SpaceAfter = 2

            For r = 1 To 5
                For c = 1 To 4

                    If r = 1 Then

                    Else
                        oTable1.Cell(r, c).Borders.Enable = True
                    End If

                    If r = 2 Then

                        If c = 1 Then
                            oTable1.Cell(r, c).Range.Font.Bold = True
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "Module"
                        ElseIf c = 2 Then
                            oTable1.Cell(r, c).Range.Font.Bold = True
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "Total"
                        ElseIf c = 3 Then
                            oTable1.Cell(r, c).Range.Font.Bold = True
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "Entered"
                        Else
                            oTable1.Cell(r, c).Range.Font.Bold = True
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "Pending"
                        End If

                    ElseIf r = 3 Then

                        If c = 1 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "Suit"
                        ElseIf c = 2 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = lpdsuit
                        ElseIf c = 3 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = lpdsuit - suit_pend
                        Else
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = suit_pend
                        End If

                    ElseIf r = 4 Then

                        If c = 1 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "RR"

                        ElseIf c = 2 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = lpdrr
                        ElseIf c = 3 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = lpdrr - rr_pend
                        Else
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = rr_pend
                        End If

                    ElseIf r = 5 Then

                        If c = 1 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "Legal action waived"

                        ElseIf c = 2 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "XX"
                        ElseIf c = 3 Then
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = legal_entered
                        Else
                            oTable1.Cell(r, c).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            oTable1.Cell(r, c).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter
                            oTable1.Cell(r, c).Range.Text = "XX"
                        End If
                    End If

                Next

            Next

            oTable1.Columns.Item(1).Width = oWord.InchesToPoints(1.7)   'Change width of columns 1 & 2
            oTable1.Columns.Item(2).Width = oWord.InchesToPoints(1.7)
            oTable1.Columns.Item(3).Width = oWord.InchesToPoints(1.7)
            oTable1.Columns.Item(4).Width = oWord.InchesToPoints(1.7)

            oTable1.Cell(1, 1).Merge(MergeTo:=oTable1.Cell(1, 4))
            oTable1.Cell(1, 1).Range.Font.Bold = True
            oTable1.Cell(1, 1).Range.Text = "Data entry status in LPD Module"

            oTable1.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable1.Cell(1, 1).Range.ParagraphFormat.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter

            oTable1.Cell(1, 1).Borders.Enable = True


            oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara4.Format.SpaceAfter = 2
            oPara4.Range.ParagraphFormat.Alignment = 3
            oPara4 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara4.Range.Text = "Branch is advised to make detailed study of the above data pertaining to NPA, LPD, Non LPD and LPD module and take immediate steps as here under:"
            oPara4.Range.Font.Name = "Calibri (Body)"   '"Abadi MT Condensed Light"
            oPara4.Range.Font.Size = 10
            oPara4.Format.SpaceAfter = 5
            oPara4.Range.InsertParagraphAfter()


            oPara4.Range.ParagraphFormat.Alignment = 3
            oPara4.Format.SpaceAfter = 2
            oPara4.Range.ListFormat.ApplyBulletDefault() 'Bullet

            oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara5.Format.SpaceAfter = 5
            oPara5.Range.ParagraphFormat.Alignment = 3

            oPara5.Range.Text = "Generate a statement by accessing NPARPT 411 and get the list of accounts which were marked as NPA prior to 01/04/2013 and is remaining without any action. Follow up each of these accounts and ensure recovery of full overdue/regularization/closure/action before 28/02/2014."
            oPara5.Range.InsertParagraphAfter()

            oPara6 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara6.Range.Text = "Verify all LPD accounts and complete the work, relating to entering the data of suit filed accounts, RR initiated accounts, Legal action waived accounts in the system by accessing the menu Suit / RR / LAW."
            oPara6.Format.SpaceAfter = 2

            oPara6.Range.InsertParagraphAfter()

            oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara5.Range.Text = "Updation of LPD module is very urgent for follow up and data generation purposes. Hence the work should be completed on a war footing basis before 15/02/2014."
            oPara5.Format.SpaceAfter = 2

            oPara5.Range.InsertParagraphAfter()

            oPara5.Range.Text = "A confirmation letter regarding completion of the above actions to be submitted to concerned RO by 01/03/2014."
            oPara5.Format.SpaceAfter = 75

            oPara5.Range.InsertParagraphAfter()
            oPara5 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara5.Range.Font.Bold = True

            oPara4.Range.ListFormat.RemoveNumbers()

            oPara5.Range.Text = "S.Radhakrishnan Nair"
            oPara5.Format.SpaceAfter = 2
            oPara5.Range.InsertParagraphAfter()
            oPara5.Range.Text = "General Manager"
            oPara5.Format.SpaceAfter = 60

            oPara5.Range.InsertParagraphAfter()

            'Page break
            oDoc.Bookmarks.Item("\endofdoc").Range.InlineShapes.AddPicture(Disk & ":\Work_assignd_17-01-2014\2.jpg")
            oRng = oDoc.Bookmarks.Item("\endofdoc").Range
            oRng.ParagraphFormat.SpaceAfter = 1
            oRng.InsertBreak(Word.WdBreakType.wdPageBreak)
            oRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

            count = count + 1
        End While
        dr.Close()


        MsgBox("Generated " & count & " pages", MsgBoxStyle.Information, "Invalid date")

    End Sub
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

            sendemail("smgbmis4@gmail.com", "ten", username, username)

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

            processmessage("Package - UPLOAD_KGB_ATM_TRAN")

            sql = "PKGEMAIL111.UPLOAD_KGB_ATM_TRAN"
            Dim cmd1 As New OracleCommand(sql, conn)
            cmd1.CommandType = CommandType.StoredProcedure
            cmd1.ExecuteNonQuery()

            processmessage("Package - DATAID_1113")

            sql = "PKGEMAIL111.DATAID_1113"
            Dim cmd4 As New OracleCommand(sql, conn)
            cmd4.CommandType = CommandType.StoredProcedure
            cmd4.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
            cmd4.ExecuteNonQuery()


            processmessage("Package - DATAID_1111")

            sql = "PKGEMAIL111.DATAID_1111"
            Dim cmd2 As New OracleCommand(sql, conn)
            cmd2.CommandType = CommandType.StoredProcedure
            cmd2.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
            cmd2.ExecuteNonQuery()

            processmessage("Package - DATAID_1112")

            sql = "PKGEMAIL111.DATAID_1112"
            Dim cmd3 As New OracleCommand(sql, conn)
            cmd3.CommandType = CommandType.StoredProcedure
            cmd3.Parameters.Add("PREVIOUSDAY", OracleDbType.Date, Nothing, ParameterDirection.Input).Value = RptDate
            cmd3.ExecuteNonQuery()

            sendemail("smgbmis4@gmail.com", "ten", username, username)

        End If
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
        Dim oracle_cnn_string As String = "Data Source=ten;C= " & username & ";Password= " & username & ";"
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
        Dim migdate As Date = InputBox("Enter Date of Migration (DD/MM/YYYY)", "Enter Value", "26-07-2014")
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

        Dim mdate As String = InputBox("Enter migration date", "", "26-07-2014")
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
        Dim mdate As String = InputBox("Enter migration date", "", "26-07-2014")
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
        If txtmenu.Text = "Aadhaar Upload - Delete Duplicate Records" Then
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
        ElseIf txtmenu.Text = "Daily emails" Then
            rptoption = 2
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder"
            lblinfo3.Text = "Ensure that outlook express is running"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Ensure that EMails are generated in the respective outboxes"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf txtmenu.Text = "Upload Files" Then
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
        ElseIf txtmenu.Text = "Tabdata" Then
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
        ElseIf txtmenu.Text = "General" Then
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
        ElseIf txtmenu.Text = "Report" Then
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
        ElseIf txtmenu.Text = "KGB Business Progress Report" Then
            rptoption = 7
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40101_XX-XX-XXXX.email as 'email.txt' and place in c:/du folder"
            lblinfo3.Text = "Rename the NMGB Trial Balance(TB_XXXXXXXX.xls) as 'nmgb.txt' (Replace tab with |) and place in C:/DU folder"
            lblinfo4.Text = "Rename the NMGB NPA(NPA_XXXXXXXX.xls) File as 'npa.txt' (Replace tab with |) and place in C:/DU folder"
            lblinfo5.Text = "Ensure that the date in all files is similar to that of previous working day"
            lblinfo6.Text = "Run the programme"
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf txtmenu.Text = "KGB Day Book" Then
            rptoption = 8
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the NMGB Trial Balance(TB_XXXXXXXX.xls) as 'nmgb.txt' (Replace tab with |) and place in C:/DU folder"
            lblinfo3.Text = "Rename the MISDO File (40124_XX-XX-XXXX.misdo) as 'smgbdb.txt' and place in C:/DU folder"
            lblinfo4.Text = "Ensure that the date in all files is similar to that of previous working day"
            lblinfo5.Text = "Run the programme"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf txtmenu.Text = "Business Review" Then
            rptoption = 9
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40102_XX-XX-XXXX.email as email2.txt and place in c:/du folder"
            lblinfo3.Text = "Enter the previous working days date in 'Previous Working Day' field"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf txtmenu.Text = "KGB First - Outstanding" Then
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
        ElseIf txtmenu.Text = "KGB First - Disbursement" Then
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
        ElseIf txtmenu.Text = "KGB First - NPA" Then
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
        ElseIf txtmenu.Text = "MISDO Upload" Then
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
        ElseIf txtmenu.Text = "ATM Data Mismatch between Finacle & Switch reports" Then
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
        ElseIf txtmenu.Text = "CIBIL Upload File Creation" Then
            rptoption = 15
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the CIBIL Individual and Non Individual files generated from Finacle in C:/DU folder"
            lblinfo3.Text = "Run the programme"
            lblinfo4.Text = "System will place the output files in D:/CIBIL folder"
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf txtmenu.Text = "EMail Daily Reports" Then
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
        ElseIf txtmenu.Text = "NPCI Linked Aadhaar - Upload file creation" Then
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
        ElseIf txtmenu.Text = "Day end EMails" Then
            rptoption = 18
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the file 40994_XX-XX-XXXX_AC1.TXT as 40994.TXT"
            lblinfo3.Text = "Rename the file 40995_XX-XX-XXXX_AC1.TXT as 40995.TXT"
            lblinfo4.Text = "Rename the file 40998AC1.TXT as 40998.TXT"
            lblinfo5.Text = "Rename the upload error file KYC_XXXXXX.TXT as KYC.TXT"
            lblinfo6.Text = "Place all files in c:/du folder"
            lblinfo7.Text = "Run the programme"
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf txtmenu.Text = "Business Review - Files to RO" Then
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
        ElseIf txtmenu.Text = "KGB Aadhar Enrolled Status" Then
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
        ElseIf txtmenu.Text = "KGB Daily Reports" Then
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
        ElseIf txtmenu.Text = "9072 Insert" Then
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
        ElseIf txtmenu.Text = "9074 Insert" Then
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
        ElseIf txtmenu.Text = "9071 Insert" Then
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
        ElseIf txtmenu.Text = "Create RO and Branch Folders and convert CIB Files" Then
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
        ElseIf txtmenu.Text = "Create Bank as a whole/All RO's/All Branches report in a single file" Then
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
        ElseIf txtmenu.Text = "Get File Names" Then
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
        ElseIf txtmenu.Text = "Word Document Generation" Then
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
        ElseIf txtmenu.Text = "Mobile Banking Transaction Status" Then
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

        ElseIf txtmenu.Text = "Create Folder" Then
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

        ElseIf txtmenu.Text = "Copy File" Then
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

        ElseIf txtmenu.Text = "Execute Script" Then
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
        ElseIf txtmenu.Text = "Basedata Generation Timing" Then
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
        ElseIf txtmenu.Text = "Staff Upload" Then
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
        ElseIf txtmenu.Text = "RO Follow Up Status Email" Then
            rptoption = 35
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename one month back email file 40101_XX-XX-XXXX.email as email_old.txt and place in c:/du folder"
            lblinfo3.Text = "Rename previousday email file 40101_XX-XX-XXXX.email as email_new.txt and place in c:/du folder"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Email will be generated in the outbox."
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf txtmenu.Text = "ATM Transaction Status" Then
            rptoption = 36
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Place the ATM transaction reports in �.txt� format in C:/DU folder."
            lblinfo3.Text = "Run the programme."
            lblinfo4.Text = "Email will be generated in the outbox."
            lblinfo5.Text = ""
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""
        ElseIf txtmenu.Text = "Inserting data into Location table" Then
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
        ElseIf txtmenu.Text = "Inserting data into CIDMASTER table" Then
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
        ElseIf txtmenu.Text = "Inserting data to Pickup table" Then
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

        ElseIf txtmenu.Text = "Inserting data to Religioncode table" Then
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
        ElseIf txtmenu.Text = "Update religioncode from banc724" Then
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
        ElseIf txtmenu.Text = "Inserting data to BranchMaster" Then

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
        ElseIf txtmenu.Text = "Inserting Deposit shadow file" Then

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
        ElseIf txtmenu.Text = "Inserting Loan shadow file" Then

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
        ElseIf txtmenu.Text = "Updating NRE code" Then

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
        ElseIf txtmenu.Text = "Inserting Staff Code" Then

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
        ElseIf txtmenu.Text = "Category code" Then

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
        ElseIf txtmenu.Text = "Inserting data to Citycode1" Then

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
        ElseIf txtmenu.Text = "Inserting data to Citycode2" Then

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
        ElseIf txtmenu.Text = "Inserting data to Minor table" Then

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
        ElseIf txtmenu.Text = "uncompress" Then

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

        ElseIf txtmenu.Text = "Inserting Param file and database" Then

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

        ElseIf txtmenu.Text = "Copying files for Creating Setup" Then

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
        ElseIf txtmenu.Text = "NRE from file" Then

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

        ElseIf txtmenu.Text = "Deceased from file" Then

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
        ElseIf txtmenu.Text = "Staff no From file" Then

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
        ElseIf txtmenu.Text = "Category from file" Then

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
        ElseIf txtmenu.Text = "Religion from file" Then

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
        ElseIf txtmenu.Text = "Handicapped from file" Then

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
        ElseIf txtmenu.Text = "LPD details from file" Then

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
        ElseIf txtmenu.Text = "Compress and email" Then

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
        ElseIf txtmenu.Text = "Differential Backup" Then

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

        ElseIf txtmenu.Text = "Upload - Extension based" Then
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
        ElseIf txtmenu.Text = "Insert into tables" Then
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
        ElseIf txtmenu.Text = "Differential Backup based on Extension" Then

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

        ElseIf txtmenu.Text = "Mirror image" Then

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

        ElseIf txtmenu.Text = "Generating CIDMaster File From dump" Then
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

        ElseIf txtmenu.Text = "Create text files in a loop" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Create Branch Data" Then
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


        ElseIf txtmenu.Text = "eNMGB Migration - Upload Migration Tool Files" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Upload CGL File" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Assign CustID and Account No" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - FUF Generation" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Reports" Then
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

        ElseIf txtmenu.Text = "Migration Tool Data Entry Status Email" Then
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

        ElseIf txtmenu.Text = "Export Oracle Data" Then
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

        ElseIf txtmenu.Text = "Backup, Drop and Import Oracle Tables" Then

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

        ElseIf txtmenu.Text = "Drop oracle user" Then

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

        ElseIf txtmenu.Text = "Figures At A Glance" Then

            rptoption = 43
            'lblinfo1.Text = "Database: Ten; User Name: EMail"
            lblinfo2.Text = "Rename the EMail file 40101_XX-XX-XXXX.email as email.txt and place in c:/du folder"
            lblinfo3.Text = "Ensure that outlook express is running"
            lblinfo4.Text = "Run the programme"
            lblinfo5.Text = "Ensure that all eMails are generated in the respective outboxes"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""

        ElseIf txtmenu.Text = "PMJDY Campaign" Then

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

        ElseIf txtmenu.Text = "Business Figures As On 30-09-2014" Then

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

        ElseIf txtmenu.Text = "Branch Intimation Letter" Then

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

        ElseIf txtmenu.Text = "eNMGB Migration - Upload 2059 Files" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Check 2059 Files" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Split CEDGE Dump" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Create History Transaction Data Dump" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Create NPA Upload Files" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Batch update of packages" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Create backup of live users" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Import Users" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Data from users" Then
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

        ElseIf txtmenu.Text = "eNMGB Migration - Zenith Backup Import" Then
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

        ElseIf txtmenu.Text = "Bulk SMS File Creation" Then
            rptoption = 46
            lblinfo1.Text = "Designation Input Options : ALL,CHAIR,GM,RM,SM,MG"
            lblinfo2.Text = "Office Type : ALL,HO,RO,BR"
            lblinfo3.Text = "Department : ALL,CW,CS,HW,IT,RL,PD"
            lblinfo4.Text = "smsupd.txt file will be generated in c:\temp folder"
            lblinfo5.Text = "Upload and confirm in live Server through BATCHUPD >> SMSBATCH"
            lblinfo6.Text = ""
            lblinfo7.Text = ""
            lblinfo8.Text = ""
            lblinfo9.Text = ""
            lblinfo10.Text = ""


        End If

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

        ' Checking whether BACOPEN,SB.TXT files exists

        'processmessage("Checking files")

        file1 = "c:\du\9106_16908570.rpt"
        'file2 = "c:\du\SB.txt"

        checkfile(file1, "Place the file in c:/du folder")
        'checkfile(file2, "Rename the file from tabdata as SB.TXT and place in c:/du folder")

        uploadfiledata(file1, username, "Y")
        'uploadfiledata(file2, username, "N")

        ' Delete existing data, if any, from c_du table

        processmessage("Deleting existing data")

        oracle_execute_non_query("ten", username, username, "truncate table z_email")

        ' Calling packages

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        processmessage("Package - PKGEMAIL113.DATAID_1135")

        sql = "PKGEMAIL113.DATAID_1135"
        Dim cmd4 As New OracleCommand(sql, conn)
        cmd4.CommandType = CommandType.StoredProcedure
        cmd4.Parameters.Add("PROCESSFLAG", OracleDbType.Varchar2, 10, Nothing, ParameterDirection.Input).Value = "BR"
        cmd4.ExecuteNonQuery()

        sendemail("smgbmis2@gmail.com", "ten", username, username)

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

        Dim sql As String
        Dim oradb As String = "Data Source=ten;User Id= " & username & ";Password= " & username & ";"
        Dim conn As New OracleConnection(oradb)
        conn.Open()
        desig = InputBox("Enter designation (ALL,GM,RM,SM,MG)", "Enter Value", "ALL")

        office = InputBox("Enter Office type(ALL,RO,HO,BR)", "Enter Value", "ALL")
        department = InputBox("Enter Department(CW,CS,HW,IT,RL,PD)", "Enter Value", "ALL")

        msg = InputBox("Enter the message")

        If desig.ToUpper = "ALL" And office.ToUpper = "ALL" And department.ToUpper = "ALL" Then
            sql = "SELECT MOBILE_NUM FROM Z_CUG WHERE OFFICE_TYPE IS NOT NULL"

        ElseIf desig.ToUpper = "ALL" And office.ToUpper = "ALL" Then

            department = "'" + department + "'"
            department = department.Replace(",", "','")
            department = department.ToUpper()

            sql = "SELECT MOBILE_NUM FROM Z_CUG WHERE OFFICE_TYPE IS NOT NULL AND DEPARTMENT IN (" + department + ")"

        ElseIf desig.ToUpper = "ALL" And department.ToUpper = "ALL" Then
            office = "'" + office + "'"
            office = office.Replace(",", "','")
            office = office.ToUpper()

            sql = "SELECT MOBILE_NUM FROM Z_CUG WHERE OFFICE_TYPE IS NOT NULL AND OFFICE_TYPE IN (" + office + ")"

        ElseIf office.ToUpper = "ALL" And department.ToUpper = "ALL" Then

            desig = "'" + desig + "'"
            desig = desig.Replace(",", "','")
            desig = desig.ToUpper()

            sql = "SELECT MOBILE_NUM FROM Z_CUG WHERE OFFICE_TYPE IS NOT NULL AND DESIGNATION IN (" + desig + ")"

        ElseIf desig.ToUpper = "ALL" Then

            department = "'" + department + "'"
            department = department.Replace(",", "','")
            department = department.ToUpper()

            office = "'" + office + "'"
            office = office.Replace(",", "','")
            office = office.ToUpper()
            sql = "SELECT MOBILE_NUM FROM Z_CUG WHERE OFFICE_TYPE IS NOT NULL AND DEPARTMENT IN (" + department + ")  AND OFFICE_TYPE IN (" + office + ")"

        ElseIf department.ToUpper = "ALL" Then

            office = "'" + office + "'"
            office = office.Replace(",", "','")
            office = office.ToUpper()

            desig = "'" + desig + "'"
            desig = desig.Replace(",", "','")
            desig = desig.ToUpper()

            sql = "SELECT MOBILE_NUM FROM Z_CUG WHERE OFFICE_TYPE IS NOT NULL AND DESIGNATION IN (" + desig + ") AND OFFICE_TYPE IN (" + office + ")"

        ElseIf office.ToUpper = "ALL" Then

            desig = "'" + desig + "'"
            desig = desig.Replace(",", "','")
            desig = desig.ToUpper()

            department = "'" + department + "'"
            department = department.Replace(",", "','")
            department = department.ToUpper()
            sql = "SELECT MOBILE_NUM FROM Z_CUG WHERE OFFICE_TYPE IS NOT NULL AND DEPARTMENT IN (" + department + ") AND DESIGNATION IN (" + desig + ") "

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
            sql = "SELECT MOBILE_NUM FROM Z_CUG WHERE OFFICE_TYPE IS NOT NULL AND DEPARTMENT IN (" + department + ") AND DESIGNATION IN (" + desig + ") AND OFFICE_TYPE IN (" + office + ")"


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
        conn.Close()
        MsgBox("File generated Successfully in C:\temp folder")
    End Sub

End Class