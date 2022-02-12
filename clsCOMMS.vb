Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports MySql.Data.MySqlClient
Imports log4net
Imports log4net.Config
Imports System.Configuration
Imports System.Reflection

Public Class clsCOMMS
    ''' <summary>
    ''' A Kiwiplan Inteface file class for .NET
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsCOMMS))
    Dim comm As New Onling.clsSQLExecutes

#Region "Local Decarations"
    Private _dbname As String
    Private _dbhost As String
    Private _Path As String
    Private _Filename As String
    Private _sb As New StringBuilder
    Private _JobNumber As String
    Private _Plant As String
    Private _FromStore As Integer
    Private _ToStore As Integer

#End Region

#Region "Constructors"
    ''' <summary>
    ''' Manual Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        _log.Info("Starting " & MethodBase.GetCurrentMethod().ToString())
    End Sub
#End Region

#Region "Private Propeeties"
    Private Property sb() As StringBuilder
        Get
            Return _sb
        End Get
        Set(ByVal value As StringBuilder)
            _sb = value
        End Set
    End Property
#End Region

#Region "Public Properties"
    Public Property FromStore() As Integer
        Get
            Return _FromStore
        End Get
        Set(ByVal value As Integer)
            _FromStore = value
        End Set
    End Property

    Public Property ToStore() As Integer
        Get
            Return _ToStore
        End Get
        Set(ByVal value As Integer)
            _ToStore = value
        End Set
    End Property

    Public Property Plant() As String
        Get
            Return _Plant
        End Get
        Set(ByVal value As String)
            _Plant = value
        End Set
    End Property

    Public Property JobNumber() As String
        Get
            Return _JobNumber
        End Get
        Set(ByVal value As String)
            _JobNumber = value
        End Set
    End Property

    Public Property Path() As String
        Get
            Return _Path
        End Get
        Set(ByVal value As String)
            _Path = value
        End Set
    End Property

    Public Property Filename() As String
        Get
            Return _Filename
        End Get
        Set(ByVal value As String)
            _Filename = value
        End Set
    End Property

    Public Property DBName() As String
        Get
            Return _dbname
        End Get
        Set(ByVal value As String)
            _dbname = value
        End Set
    End Property

    Public Property DBHost() As String
        Get
            Return _dbhost
        End Get
        Set(ByVal value As String)
            _dbhost = value
        End Set
    End Property
#End Region

#Region "Private Functions"
    ''' <summary>
    ''' Checks a field if Numeric
    ''' </summary>
    ''' <param name="number">Pass any object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsNumeric(ByVal number As Object) As Boolean
        Dim i As Integer
        For i = 0 To number.Length - 1
            If Not Char.IsNumber(number, i) Then
                Return False
            End If
        Next
        Return True
    End Function
    ''' <summary>
    ''' Checks a field if Alpha
    ''' </summary>
    ''' <param name="str">Pass any object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsChar(ByVal str As Object) As Boolean
        Dim i As Integer
        For i = 0 To str.Length - 1
            If Not Char.IsLetter(str, i) Then
                Return False
            End If
        Next
        Return True
    End Function

#End Region

#Region "Public Functions"
    ''' <summary>
    ''' Function to Build the COMMS file
    ''' </summary>
    ''' <param name="Exported">File Exported Previously True of False</param>
    ''' <returns>Returns the file in a string builder fasion</returns>
    ''' <remarks></remarks>
    Public Function BuildFile(ByVal Exported As Boolean)
        Try

            'Credentials for Sql
            comm.SqlDBName = _dbname
            comm.SqlHostname = _dbhost
            comm.SqlConn(True)

            'Was the file Exported before
            If Not Exported Then
                CSDATA1(False)
                CSDATA2(False)
                CSDATA3(False)
                ADOD1(False)
                ADOD2(False)
                ADOD3(False)
                ADPCS(False)
            Else
                CSDATA1(True)
                CSDATA2(True)
                CSDATA3(True)
                ADOD1(True)
                ADOD2(True)
                ADOD3(True)
                ADPCS(True)
            End If

            'Create File on disk
            WriteFile()

        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create build the file..")
        Finally

        End Try
        Return sb
    End Function
#End Region

#Region "Create ADD Record"

#End Region

#Region "Create LDAD1 Record"
    Private Sub CSDATA1(ByVal Exported As Boolean)
        Try
            'Get the Data needed to start building the record
            Dim strSQL = "select  " & _
                            "o.Jobnumber, " & _
                            "c.Name," & _
                            "a.Street, " & _
                            "a.City, " & _
                            "a.Country, " & _
                            "a.DestinationCode, " & _
                            "p.CustomerSpec, " & _
                            "p.Description, " & _
                            "s.PalletType, " & _
                            "s.QuantityPerBundle, " & _
                            "s.BundlesPerLayer, " & _
                            "s.LayersPerUnit, " & _
                            "s.LengthWiseStrapsPerUnit, " & _
                            "s.WidthWiseStrapsPerUnit, " & _
                            "s.StackingPattern, " & _
                            "s.UnitCovering, " & _
                            "o.CustomerPORef, " & _
                            "p.DesignNumber, " & _
                            "s.LabelsPerUnit, " & _
                            "StrappingCode='1', " & _
                            "s.PiecesPerUnit, " & _
                            "s.StrapperCompression, " & _
                            "s.RotateBefore, " & _
                            "s.RotateAfter, " & _
                            "s.EdgeProtector, " & _
                            "s.RotateOffCorrugator, " & _
                            "s.RotateBefore, " & _
                            "s.CSCLabelsPerUnit " & _
                            "From esporder o " & _
                            "INNER Join orgcompany c " & _
                            "ON c.id = o.companyid " & _
                            "INNER Join orgaddress a " & _
                            "ON a.id = o.shiptoaddressID " & _
                            "INNER Join ebxproductdesign p " & _
                            "ON p.id = o.productdesignID " & _
                            "INNER Join espunitizingspec s " & _
                            "on o.companyid = c.id and o.jobnumber = s.jobnumber " & _
                            "where o.jobnumber='" & _JobNumber & "'"

            Dim objDS As DataSet = comm.FillDatasetSQL(strSQL, CommandType.Text, Nothing)

            Dim dr As DataRow
            'Loop throw every row
            For Each dr In objDS.Tables(0).Rows

                ''Create the Container
                'Dim sb As New StringBuilder

                'Was the file Exported before
                If Not Exported Then
                    sb.Append("LDAD1") 'Transaction request
                Else
                    sb.Append("LDCH1") 'Transaction request
                End If
                sb.Append(" ".ToString.PadRight(5, " ")) 'Transaction answer
                sb.Append(" ".ToString.PadRight(2, " ")) 'CSDATA.printed_flag
                sb.Append(Mid(dr.Item("jobnumber").ToString, 1, 10).PadRight(10, " ")) 'CSDATA.job_key
                sb.Append(Mid(dr.Item("Name").ToString, 1, 32).PadRight(32, " ")) 'CSDATA.full_customer_name
                sb.Append(Mid(dr.Item("Street").ToString, 1, 32).PadRight(32, " ")) 'CSDATA.shipto_address_1
                sb.Append(Mid(dr.Item("City").ToString, 1, 20).PadRight(20, " ")) 'CSDATA.shipto_address_2
                sb.Append(Mid(dr.Item("Country").ToString, 1, 20).PadRight(20, " ")) 'CSDATA.shipto_address_3
                sb.Append(Mid(dr.Item("DestinationCode").ToString, 1, 10).PadRight(10, " ")) 'CSDATA.delivery_code
                sb.Append(Mid(dr.Item("CustomerSpec").ToString, 1, 30).PadRight(30, " ")) 'CSDATA.product_desc_1
                sb.Append(Mid(dr.Item("Description").ToString, 1, 30).PadRight(30, " ")) 'CSDATA.product_desc_2
                sb.Append(Mid(dr.Item("PalletType").ToString, 1, 2).PadRight(2, " ")) 'CSDATA.pallet_type
                sb.Append(Mid(dr.Item("QuantityPerBundle").ToString, 1, 4).PadLeft(4, "0")) 'CSDATA.qty_per_bundle
                sb.Append(Mid(dr.Item("BundlesPerLayer").ToString, 1, 3).PadLeft(3, "0")) 'CSDATA.bundles_layer
                sb.Append("".PadRight(1, " ")) 'spare
                sb.Append(Mid(dr.Item("LayersPerUnit").ToString, 1, 3).PadLeft(3, "0")) 'CSDATA.layers_per_unit
                sb.Append("".PadRight(6, " ")) 'reserved
                sb.Append(Mid(dr.Item("LengthWiseStrapsPerUnit").ToString, 1, 3).PadLeft(3, "0")) 'CSDATA.straps_lengthwise
                sb.Append(Mid(dr.Item("WidthWiseStrapsPerUnit").ToString, 1, 3).PadLeft(3, "0")) 'CSDATA.straps_widthwise
                sb.Append("".PadRight(1, " ")) 'spare
                sb.Append(Mid(dr.Item("StackingPattern").ToString, 1, 6).PadRight(6, " ")) 'CSDATA.stack_pattern
                sb.Append(Mid(dr.Item("UnitCovering").ToString, 1, 12).PadRight(12, " ")) 'CSDATA.unit_covering
                sb.Append(Mid(dr.Item("CustomerPORef").ToString, 1, 25).PadRight(25, " ")) 'CSDATA.customer_order_num
                sb.Append(Mid(dr.Item("DesignNumber").ToString, 1, 25).PadRight(25, " ")) 'CSDATA.specification
                sb.Append("".PadRight(10, " ")) 'spare
                sb.Append(Mid(dr.Item("LabelsPerUnit").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.fgs_tags_pallet
                sb.Append(Mid(dr.Item("StrappingCode").ToString, 1, 4).PadRight(4, " ")) 'CSDATA.strap_code
                sb.Append(Mid(dr.Item("PiecesPerUnit").ToString, 1, 6).PadLeft(6, "0")) 'CSDATA.quantity_per_unit
                sb.Append(Mid(dr.Item("StrapperCompression").ToString, 1, 1).PadRight(1, " ")) 'CSDATA.platen_compression
                sb.Append(Mid(dr.Item("RotateBefore").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.rotate_on_entry
                sb.Append(Mid(dr.Item("RotateAfter").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.rotate_on_exit
                sb.Append(Mid(dr.Item("EdgeProtector").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.edge_protect
                sb.Append(Mid(dr.Item("RotateOffCorrugator").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.off_corr_rotate
                sb.Append(Mid(dr.Item("RotateBefore").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.pre_conv_rotate
                sb.Append(Mid(dr.Item("CSCLabelsPerUnit").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.labels_per_unit
                sb.Append("".PadRight(1, " ")) 'EndFeed
                sb.Append(vbCr) 'End

            Next
            _log.Info("Created CSDATA 1: ")
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create the CSDATA 1 record..")
        Finally
            _log.Info("Created the CSDATA 1 record..")

        End Try
    End Sub
#End Region

#Region "Create LDAD2 Record"
    Private Sub CSDATA2(ByVal Exported As Boolean)
        Try
            'Get the Data needed to start building the record
            Dim strSQL = "select  " & _
                            "o.Jobnumber, " & _
                            "UserDefined ='', " & _
                            "s.WidthWiseStacksPerUnit, " & _
                            "s.LengthWiseStacksPerUnit, " & _
                            "CASE isnull(s.pallettype,'') When '' then 0 else 1 end WidthWisePalletPerUnit, " & _
                            "CASE isnull(s.pallettype,'') When '' then 0 else 1 end LengthWisePalletPerUnit, " & _
                            "s.ReverseStack, " & _
                            "s.TopBoardCode, " & _
                            "s.MaxUnitWidth, " & _
                            "s.MaxUnitLength, " & _
                            "s.MaxStackHeight, " & _
                            "CASE isnull(a.CorrugatorInHouseLabel,'') When '' then isnull(p.CorrugatorInHouseLabel,'') End CorrugatorInHouseLabel, " & _
                            "isnull(a.CorrugatorDeliveryLabel,'') CorrugatorDeliveryLabel, " & _
                            "CASE isnull(a.ConvertingInHouseLabel,'') When '' then isnull(p.ConvertingInHouseLabel,'') End ConvertingInHouseLabel, " & _
                            "CASE isnull(a.ConvertingDeliveryLabel,'') When '' then isnull(u.LabelFormat,'') End ConvertingDeliveryLabel, " & _
                            "Dummy1='', " & _
                            "Dummy2='', " & _
                            "u.UnitHeight " & _
                            "From esporder o " & _
                            "INNER Join orgcompany c " & _
                            "ON c.id = o.companyid " & _
                            "INNER Join orgaddress a " & _
                            "ON a.id = o.shiptoaddressID " & _
                            "INNER Join ebxproductdesign p " & _
                            "ON p.id = o.productdesignID " & _
                            "INNER Join espunitizingspec s " & _
                            "on o.companyid = c.id and o.jobnumber = s.jobnumber " & _
                            "INNER Join ebxunitizingdata u " & _
                            "ON p.unitizingdataID = u.id " & _
                            "where o.jobnumber='" & _JobNumber & "'"

            Dim objDS As DataSet = comm.FillDatasetSQL(strSQL, CommandType.Text, Nothing)

            Dim dr As DataRow
            'Loop throw every row
            For Each dr In objDS.Tables(0).Rows

                'Was the file Exported before
                If Not Exported Then
                    sb.Append("LDAD2") 'Transaction request
                Else
                    sb.Append("LDCH2") 'Transaction request
                End If
                sb.Append("".ToString.PadRight(5, " ")) 'Transaction answer
                sb.Append("".ToString.PadRight(2, " ")) 'CSDATA.printed_flag
                sb.Append(Mid(dr.Item("jobnumber").ToString, 1, 10).PadRight(10, " ")) 'CSDATA.job_key
                sb.Append("".PadRight(6, " ")) 'spare

                '=======
                'This may need another function prior creating this line to get all add info
                sb.Append(Mid(dr.Item("UserDefined").ToString, 1, 240).PadRight(240, " ")) 'CSDATA.user_defined
                '=======

                sb.Append(Mid(dr.Item("WidthWiseStacksPerUnit").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.stacks_widthwise
                sb.Append(Mid(dr.Item("LengthWiseStacksPerUnit").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.stacks_lenghtwise
                sb.Append(Mid(dr.Item("WidthWisePalletPerUnit").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.pallets_widthwise
                sb.Append(Mid(dr.Item("LengthWisePalletPerUnit").ToString, 1, 1).PadLeft(1, "0")) 'CSDATA.pallets_lengthwise
                sb.Append(Mid(dr.Item("ReverseStack").ToString, 1, 1).PadRight(1, " ")) 'CSDATA.reverse_stack
                sb.Append(Mid(dr.Item("TopBoardCode").ToString, 1, 3).PadRight(3, " ")) 'CSDATA.top_board_code
                sb.Append(Mid(dr.Item("MaxUnitWidth").ToString, 1, 6).PadLeft(6, "0")) 'CSDATA.max_unit_width
                sb.Append(Mid(dr.Item("MaxUnitLength").ToString, 1, 6).PadLeft(6, "0")) 'CSDATA.max_unit_length
                sb.Append(Mid(dr.Item("MaxStackHeight").ToString, 1, 6).PadLeft(6, "0")) 'CSDATA.max_Stack_height
                sb.Append(Mid(dr.Item("CorrugatorInHouseLabel").ToString, 1, 2).PadRight(2, " ")) 'CSDATA.label_csc.wip
                sb.Append(Mid(dr.Item("CorrugatorDeliveryLabel").ToString, 1, 2).PadRight(2, " ")) 'CSDATA.label_csc.wip
                sb.Append(Mid(dr.Item("ConvertingInHouseLabel").ToString, 1, 2).PadRight(2, " ")) 'CSDATA.label_wip_wip
                sb.Append(Mid(dr.Item("ConvertingDeliveryLabel").ToString, 1, 2).PadRight(2, " ")) 'CSDATA.label_wip_fgs
                sb.Append(Mid(dr.Item("Dummy1").ToString, 1, 2).PadRight(2, " ")) 'CSDATA.label_stk_wip
                sb.Append(Mid(dr.Item("Dummy2").ToString, 1, 2).PadRight(2, " ")) 'CSDATA.label_stk_fgs
                sb.Append(Mid(dr.Item("UnitHeight").ToString, 1, 6).PadLeft(6, "0")) 'CSDATA.unit_height
                sb.Append("".PadRight(6, " ")) 'spare
                sb.Append("".PadRight(1, " ")) 'EndFeed
                sb.Append(vbCr) 'End

            Next
            _log.Info("Created CSDATA 2: ")
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create the CSDATA 2 record..")
        Finally
            _log.Info("Created the CSDATA 2 record..")

        End Try
    End Sub
#End Region

#Region "Create LDAD3 Record"
    Private Sub CSDATA3(ByVal Exported As Boolean)
        Try
            'Get the Data needed to start building the record
            Dim strSQL = "select  " & _
                            "o.Jobnumber, " & _
                            "Shipper = '' " & _
                            "From esporder o " & _
                            "INNER Join orgcompany c " & _
                            "ON c.id = o.companyid " & _
                            "INNER Join orgaddress a " & _
                            "ON a.id = o.shiptoaddressID " & _
                            "INNER Join ebxproductdesign p " & _
                            "ON p.id = o.productdesignID " & _
                            "INNER Join espunitizingspec s " & _
                            "on o.companyid = c.id and o.jobnumber = s.jobnumber " & _
                            "INNER Join ebxunitizingdata u " & _
                            "ON p.unitizingdataID = u.id " & _
                            "where o.jobnumber='" & _JobNumber & "'"

            Dim objDS As DataSet = comm.FillDatasetSQL(strSQL, CommandType.Text, Nothing)

            Dim dr As DataRow
            'Loop throw every row
            For Each dr In objDS.Tables(0).Rows

                'Was the file Exported before
                If Not Exported Then
                    sb.Append("LDAD3") 'Transaction request
                Else
                    sb.Append("LDCH3") 'Transaction request
                End If
                sb.Append("".ToString.PadRight(5, " ")) 'Transaction answer
                sb.Append("".ToString.PadRight(2, " ")) 'spare
                sb.Append(Mid(dr.Item("jobnumber").ToString, 1, 10).PadRight(10, " ")) 'CSDATA.job_key
                sb.Append(Mid(dr.Item("Shipper").ToString, 1, 256).PadRight(296, " ")) 'spare
                sb.Append("".PadRight(1, " ")) 'EndFeed
                sb.Append(vbCr) 'End

            Next
            _log.Info("Created CSDATA 3: ")
            _log.Info(sb.ToString)
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create the CSDATA 3 record..")
        Finally
            _log.Info("Created the CSDATA 3 record..")

        End Try
    End Sub
#End Region

#Region "Create ADOD1 Record"
    Private Sub ADOD1(ByVal Exported As Boolean)
        Try
            'Get the Data needed to start building the record
            Dim strSQL = "select  " & _
                            "o.JobNumber, " & _
                            "p.DesignNumber, " & _
                            "o.CustomerPORef, " & _
                            "c.CompanyNumber, " & _
                            "c.Name, " & _
                            "o.OrderedQuantity, " & _
                            "r.startlength, " & _
                            "r.startwidth, " & _
                            "p.FinishedLength, " & _
                            "p.FinishedWidth, " & _
                            "Scores ='', " & _
                            "ScoreCodes ='', " & _
                            "GlueFlapCode='F', " & _
                            "Steps=2, " & _
                            "p.SumtoWidthSlotDepth01, " & _
                            "p.SumtoWidthSlotDepth02, " & _
                            "p.SumtoLengthSlotDepth01, " & _
                            "p.SumtoLengthSlotDepth02, " & _
                            "DifficultyFactor=1, " & _
                            "isnull(p.HoleType,'') HoleType, " & _
                            "p.ClosureCode, " & _
                            "st.StyleCode, " & _
                            "JobStatus=1, " & _
                            "substring(isnull(ct.firstname,''),1,1) + '' + isnull(ct.lastname,'') RepContact, " & _
                            "s.QuantityPerBundle, " & _
                            "Round(o.GoodsValue,0) GoodsValue, " & _
                            "o.wipvalue, " & _
                            "isnull(o.Priority,'') Priority, " & _
                            "Export='N', " & _
                            "r.KnockOutWaste, " & _
                            "r.Knife, " & _
                            "r.printingplate, " & _
                            "PrintQuality='', " & _
                            "PrintStatus='', " & _
                            "KnifeStatus='', " & _
                            "o.OrderType " & _
                            "From esporder o " & _
                            "INNER Join ebxroute r " & _
                            "ON r.id = o.routeid " & _
                            "INNER Join orgcompany c " & _
                            "ON c.id = o.companyid " & _
                            "INNER Join orgaddress a " & _
                            "ON a.id = o.shiptoaddressID " & _
                            "INNER Join ebxproductdesign p " & _
                            "ON p.id = o.productdesignID " & _
                            "INNER Join ebxstyle st " & _
                            "ON st.id = p.StyleID " & _
                            "INNER Join espunitizingspec s " & _
                            "on o.companyid = c.id and o.jobnumber = s.jobnumber " & _
                            "INNER Join ebxunitizingdata u " & _
                            "ON p.unitizingdataID = u.id " & _
                            "LEFT Outer Join orgcontact ct " & _
                            "ON ct.id = a.repcontactID " & _
                            "where o.jobnumber='" & _JobNumber & "'"

            Dim objDS As DataSet = comm.FillDatasetSQL(strSQL, CommandType.Text, Nothing)

            Dim dr As DataRow
            'Loop throw every row
            For Each dr In objDS.Tables(0).Rows

                'Was the file Exported before
                If Not Exported Then
                    sb.Append("ADOD1") 'Transaction request
                Else
                    sb.Append("CHOD1") 'Transaction request
                End If
                sb.Append("".ToString.PadRight(5, " ")) 'Transaction answer
                sb.Append(Mid(dr.Item("JobNumber").ToString, 1, 10).PadRight(10, " ")) 'JBSPEC.job_number
                sb.Append(Mid(dr.Item("DesignNumber").ToString, 1, 10).PadRight(10, " ")) 'JBSPEC.spec_number
                sb.Append(Mid(dr.Item("CustomerPORef").ToString, 1, 25).PadRight(25, " ")) 'JBSPEC.customer_order_num
                sb.Append(Mid(dr.Item("CompanyNumber").ToString, 1, 10).PadRight(10, " ")) 'JBSPEC.customer_number
                sb.Append(Mid(Mid(dr.Item("Name").ToString, 1, 10), 1, 10).PadRight(10, " ")) 'JBSPEC.customer_name
                sb.Append(Mid(dr.Item("OrderedQuantity").ToString, 1, 8).PadLeft(8, "0")) 'JBSPEC.customer_order_num
                sb.Append(Mid(dr.Item("StartLength").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.initial_length
                sb.Append(Mid(dr.Item("StartWidth").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.initial_width
                sb.Append(Mid(dr.Item("FinishedLength").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.final_length
                sb.Append(Mid(dr.Item("FinishedWidth").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.final_width
                sb.Append(Mid(dr.Item("Scores").ToString, 1, 100).PadLeft(100, "0")) 'JBSPEC.score -> Not Sending
                sb.Append(Mid(dr.Item("ScoreCodes").ToString, 1, 24).PadRight(24, " ")) 'JBSPEC.score_codes -> Not Sending
                sb.Append(Mid(dr.Item("GlueFlapCode").ToString, 1, 1).PadRight(1, " ")) 'JBSPEC.gleu_flap_code
                sb.Append(Mid(dr.Item("Steps").ToString, 1, 2).PadLeft(2, "0")) 'JBSPEC.number_of_steps
                sb.Append("".ToString.PadRight(1, " ")) 'spare
                sb.Append(Mid(dr.Item("SumtoWidthSlotDepth01").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.slot_depth_corr_1
                sb.Append(Mid(dr.Item("SumtoWidthSlotDepth02").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.slot_depth_corr_2
                sb.Append(Mid(dr.Item("SumtoLengthSlotDepth01").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.slot_depth_across1
                sb.Append(Mid(dr.Item("SumtoLengthSlotDepth02").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.slot_depth_across2
                sb.Append(Mid(dr.Item("DifficultyFactor").ToString, 1, 1).PadLeft(1, "0")) 'JBSPEC.difficulty_factor
                sb.Append(Mid(dr.Item("HoleType").ToString, 1, 1).PadRight(1, " ")) 'JBSPEC.sauer_equipment
                sb.Append(Mid(dr.Item("ClosureCode").ToString, 1, 2).PadRight(2, " ")) 'JBSPEC.closure_code
                sb.Append(Mid(dr.Item("StyleCode").ToString, 1, 8).PadRight(8, " ")) 'JBSPEC.fefco_code
                sb.Append(Mid(dr.Item("JobStatus").ToString, 1, 1).PadLeft(1, "0")) 'JBSPEC.job_status
                sb.Append(Mid(dr.Item("RepContact").ToString, 1, 4).PadRight(4, " ")) 'JBSPEC.sale_rep_code
                sb.Append(Mid(dr.Item("QuantityPerBundle").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.unit_quantity
                sb.Append(Mid(dr.Item("GoodsValue").ToString, 1, 10).PadLeft(10, "0").Replace(".", "0")) 'JBSPEC.selling_price
                sb.Append(Mid(dr.Item("WipValue").ToString, 1, 10).PadLeft(10, "0").Replace(".", "0")) 'JBSPEC.valuation
                sb.Append(Mid(dr.Item("Priority").ToString, 1, 2).PadLeft(2, "0")) 'JBSPEC.priority
                sb.Append(Mid(dr.Item("Export").ToString, 1, 1).PadRight(1, " ")) 'JBSPEC.esport_customer
                sb.Append(Mid(dr.Item("KnockOutWaste").ToString, 1, 4).PadLeft(4, "0")) 'STEDIE.die_waste_percent
                sb.Append(Mid(dr.Item("Knife").ToString, 1, 10).PadRight(10, " ")) 'JBSPEC.die_number
                sb.Append(Mid(dr.Item("PrintingPlate").ToString, 1, 10).PadRight(10, " ")) 'JBSPEC.print_number
                sb.Append("".ToString.PadRight(1, " ")) 'spare
                sb.Append(Mid(dr.Item("PrintQuality").ToString, 1, 1).PadLeft(1, "0")) 'JBSPEC.print_quality
                sb.Append(Mid(dr.Item("PrintQuality").ToString, 1, 1).PadLeft(2, "0")) 'STEDIE.tool_status (die)
                sb.Append(Mid(dr.Item("PrintQuality").ToString, 1, 2).PadLeft(2, "0")) 'STEDIE.tool_status (print)
                sb.Append(Mid(dr.Item("OrderType").ToString, 1, 1).PadLeft(1, "0")) 'JBSPEC.stock_manuf_flag
                sb.Append("".PadRight(1, " ")) 'EndFeed
                sb.Append(vbCr) 'End

            Next
            _log.Info("Created LDAD 1: ")
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create the LDAD 1 record..")
        Finally
            _log.Info("Created the LDAD 1 record..")

        End Try
    End Sub
#End Region

#Region "Create ADOD2 Record"
    Private Sub ADOD2(ByVal Exported As Boolean)
        Try
            'Get the Data needed to start building the record
            Dim strSQL = "select  " & _
                            "o.Jobnumber, " & _
                            "CASE when isnull(pd.DesignNumber,0) =0 then p.designnumber end OldSpecNumber, " & _
                            "Inks='', " & _
                            "a.AddressNumber, " & _
                            "s.PalletType, " & _
                            "AssemblyTime='', " & _
                            "a.PreferredDeliveryTime, " & _
                            "o.DueDate, " & _
                            "fz.JourneyDistance, " & _
                            "CASE when isnull(dc.noperset,0) = 0 then p.noperset else dc.noperset end NoPerSet, " & _
                            "CASE when isnull(dc.ComponentNo,0) = 0 then 1 else dc.ComponentNo end ComponentNo, " & _
                            "CASE when isnull(dc.partproductdesignID,0) = 0 then p.noperset else (Select Count(*) from ebxdesigncomponent where parentproductdesignid=dc.parentproductdesignID) end NoOfParts, " & _
                            "Scores='', " & _
                            "ScoreCodes='', " & _
                            "o.PermittedUnderrun, " & _
                            "o.PermittedOverrun, " & _
                            "a.DestinationCode, " & _
                            "p.InternalLength, " & _
                            "p.InternalWidth, " & _
                            "p.InternalDepth, " & _
                            "p.AdditionalClosureCode, " & _
                            "o.DueDate, " & _
                            "p.Boardcode, " & _
                            "GlueFlapCode='F', " & _
                            "isnull(p.CustomerSpec,'') CustomerSpec, " & _
                            "p.Description, " & _
                            "o.EarliestDuedate, " & _
                            "SlotCode='', " & _
                            "substring(ct.firstname,1,1)+''+ct.lastname SalesRep, " & _
                            "PercentOverOrder='', " & _
                            "TruckUnitLoad='', " & _
                            "a.DespatchMode, " & _
                            "o.OrderType " & _
                            "From esporder o " & _
                            "INNER Join orgcompany c " & _
                            "ON c.id = o.companyid " & _
                            "INNER Join orgaddress a " & _
                            "ON a.id = o.shiptoaddressID " & _
                            "INNER Join orgdeliveryregion dr " & _
                            "ON a.deliveryregionID = dr.id " & _
                            "INNER JOIN orgfreightzone fz " & _
                            "on fz.deliveryregionid = dr.id " & _
                            "INNER Join ebxproductdesign p " & _
                            "LEFT Outer Join ebxproductdesign pd " & _
                            "ON pd.id = p.usehistoryfromproductdesignid " & _
                            "LEFT Outer Join ebxdesigncomponent dc " & _
                            "ON p.id = dc.partproductdesignID " & _
                            "ON p.id = o.productdesignID " & _
                            "INNER Join espunitizingspec s " & _
                            "on o.companyid = c.id and o.jobnumber = s.jobnumber " & _
                            "INNER Join ebxunitizingdata u " & _
                            "ON p.unitizingdataID = u.id " & _
                            "LEFT Outer Join orgcontact ct " & _
                            "ON ct.id = a.SalescontactID " & _
                            "where o.jobnumber='" & _JobNumber & "'"

            Dim objDS As DataSet = comm.FillDatasetSQL(strSQL, CommandType.Text, Nothing)

            Dim dr As DataRow
            'Loop throw every row
            For Each dr In objDS.Tables(0).Rows

                'Was the file Exported before
                If Not Exported Then
                    sb.Append("ADOD2") 'Transaction request
                Else
                    sb.Append("CHOD2") 'Transaction request
                End If
                sb.Append("".ToString.PadRight(5, " ")) 'Transaction answer
                sb.Append(Mid(dr.Item("jobnumber").ToString, 1, 10).PadRight(10, " ")) 'JBSPEC.job_number
                sb.Append(Mid(dr.Item("OldSpecNumber").ToString, 1, 10).PadRight(10, " ")) 'JBSPEC.old_spec_number
                sb.Append(Mid(dr.Item("Inks").ToString, 1, 48).PadRight(48, " ")) 'JBSPEC.ink_code -> Not Sending
                sb.Append(Mid(dr.Item("AddressNumber").ToString, 1, 3).PadLeft(3, "0")) 'JBSPEC.address_number
                sb.Append(Mid(dr.Item("PalletType").ToString, 1, 2).PadRight(2, " ")) 'JBSPEC.unit_load_code
                sb.Append(Mid(dr.Item("AssemblyTime").ToString, 1, 4).PadRight(4, " ")) 'JBSPEC.assembly_time ->Set To Blank but Numeric
                sb.Append(Mid(Format(dr.Item("PreferredDeliveryTime"), "hhmm").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.prefer_delivertime
                sb.Append(Mid(Format(dr.Item("DueDate"), "hhmm").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.delivery_time
                sb.Append(Mid(dr.Item("JourneyDistance").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.delivery_distance
                sb.Append(Mid(dr.Item("NoPerSet").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.number_per_set
                sb.Append(Mid(dr.Item("ComponentNo").ToString, 1, 2).PadLeft(2, "0")) 'JBSPEC.part_number
                sb.Append(Mid(dr.Item("NoOfParts").ToString, 1, 2).PadLeft(2, "0")) 'JBSPEC.total_number_parts
                sb.Append(Mid(dr.Item("Scores").ToString, 1, 40).PadRight(40, " ")) 'JBSPEC.scores_across
                sb.Append(Mid(dr.Item("ScoreCodes").ToString, 1, 9).PadRight(9, " ")) 'JBSPEC.score_across_cod
                sb.Append(Mid(dr.Item("PermittedUnderrun").ToString, 1, 2).PadRight(2, " ")) 'JBSPEC.underrun_percent
                sb.Append(Mid(dr.Item("PermittedOverrun").ToString, 1, 2).PadRight(2, " ")) 'JBSPEC.overrun_percent
                sb.Append(Mid(dr.Item("DestinationCode").ToString, 1, 4).PadRight(4, " ")) 'JBSPEC.destination_code
                sb.Append(Mid(dr.Item("InternalLength").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.internal_length
                sb.Append(Mid(dr.Item("InternalWidth").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.internal_width
                sb.Append(Mid(dr.Item("InternalDepth").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.internal_height
                sb.Append(Mid(dr.Item("AdditionalClosureCode").ToString, 1, 2).PadRight(2, " ")) 'JBSPEC.extra_joint_code
                sb.Append("".PadLeft(2, "0")) 'spare
                sb.Append(Mid(Format(dr.Item("DueDate"), "yyMMdd").ToString, 1, 6).PadLeft(6, "0")) 'JBSPEC.due_date
                sb.Append(Mid(Format(dr.Item("DueDate"), "hhmm").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.due_tim
                sb.Append(Mid(dr.Item("BoardCode").ToString, 1, 30).PadRight(30, " ")) 'Board type
                sb.Append(Mid(dr.Item("GlueFlapCode").ToString, 1, 1).PadRight(1, " ")) 'JBSPEC.glue_flap_across
                sb.Append(Mid(dr.Item("CustomerSpec").ToString, 1, 30).PadRight(30, " ")) 'JBSPEC.job_description
                sb.Append(Mid(dr.Item("Description").ToString, 1, 24).PadRight(24, " ")) 'JBSPEC.productdescription
                sb.Append(Mid(Format(dr.Item("EarliestDueDate"), "yyMMdd").ToString, 1, 6).PadLeft(6, "0")) 'JBSPEC.early_dt
                sb.Append(Mid(Format(dr.Item("EarliestDuedate"), "hhmm").ToString, 1, 4).PadLeft(4, "0")) 'JBSPEC.early_tm
                sb.Append(Mid(dr.Item("SlotCode").ToString, 1, 1).PadRight(1, " ")) 'JBSPEC.slot_code -> Not Sending
                sb.Append(Mid(dr.Item("SalesRep").ToString, 1, 4).PadRight(4, " ")) 'JBSPEC.sales_clerk_code
                sb.Append(Mid(dr.Item("PercentOverOrder").ToString, 1, 2).PadRight(2, " ")) 'JBSPEC.percent_over_order -> Not Sending but Numeric
                sb.Append(Mid(dr.Item("TruckUnitLoad").ToString, 1, 8).PadRight(8, " ")) 'JBSPEC.truck_unit_load -> Not Sending but Numeric
                sb.Append(Mid(dr.Item("DespatchMode").ToString, 1, 8).PadRight(8, " ")) 'JBSPEC.shipment_method
                sb.Append("".PadRight(9, " ")) 'spare
                sb.Append(Mid(dr.Item("OrderType").ToString, 1, 1).PadRight(1, " ")) 'JBSPEC.stock_manuf_flag
                sb.Append("".PadRight(1, " ")) 'EndFeed
                sb.Append(vbCr) 'End

            Next
            _log.Info("Created ADOD 2: ")
            _log.Info(sb.ToString)
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create the ADOD 2 record..")
        Finally
            _log.Info("Created the ADOD 2 record..")

        End Try
    End Sub
#End Region

#Region "Create ADOD3 Record"
    Private Sub ADOD3(ByVal Exported As Boolean)
        Try
            'Get the Data needed to start building the record
            Dim strSQL = "select  " & _
                            "o.Jobnumber " & _
                            "From esporder o " & _
                            "INNER Join orgcompany c " & _
                            "ON c.id = o.companyid " & _
                            "INNER Join orgaddress a " & _
                            "ON a.id = o.shiptoaddressID " & _
                            "INNER Join ebxproductdesign p " & _
                            "ON p.id = o.productdesignID " & _
                            "INNER Join espunitizingspec s " & _
                            "on o.companyid = c.id and o.jobnumber = s.jobnumber " & _
                            "INNER Join ebxunitizingdata u " & _
                            "ON p.unitizingdataID = u.id " & _
                            "where o.jobnumber='" & _JobNumber & "'"

            Dim objDS As DataSet = comm.FillDatasetSQL(strSQL, CommandType.Text, Nothing)

            Dim dr As DataRow
            'Loop throw every row
            For Each dr In objDS.Tables(0).Rows

                'Was the file Exported before
                If Not Exported Then
                    sb.Append("ADOD3") 'Transaction request
                Else
                    sb.Append("CHOD3") 'Transaction request
                End If
                sb.Append("".ToString.PadRight(5, " ")) 'Transaction answer
                sb.Append(dr.Item("jobnumber").ToString.PadRight(10, " ")) 'JBSPEC.job_number
                sb.Append("".PadRight(298, " ")) 'Dummy
                sb.Append("".PadRight(1, " ")) 'EndFeed
                sb.Append(vbCr) 'End

            Next
            _log.Info("Created ADOD 3: ")
            _log.Info(sb.ToString)
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create the ADOD 3 record..")
        Finally
            _log.Info("Created the ADOD 3 record..")

        End Try
    End Sub
#End Region

#Region "Create ADPCS Record"
    Private Sub ADPCS(ByVal Exported As Boolean)
        Try
            'Get the machine needed for the Storage step
            Dim strSQLM = "select pv.value " & _
                            "from infplantvalue pv " & _
                            "INNER Join infparameter p " & _
                            "ON p.id = pv.parameterid " & _
                            "INNER Join orgplant pl " & _
                            "ON pl.id = pv.plantID " & _
                            "where p.name like 'Store '+'" & _ToStore & "'+' - In Machine%' " & _
                            "and pl.plantnumber='" & _Plant & "'"

            Dim objDSM As DataSet = comm.FillDatasetSQL(strSQLM, CommandType.Text, Nothing)

            'Set the Machine to a variable
            Dim _Machine = objDSM.Tables(0).Rows.Item(0).Item("value").ToString

            'Get the Data needed to start building the record
            Dim strSQL = "select  " & _
                            "o.Jobnumber, " & _
                            "SeriesNumber='1', " & _
                            "StepNumber='2', " & _
                            "EntryNumberUP1='1', " & _
                            "EntryNumberUP2='1', " & _
                            "ExitNumberUP1='1', " & _
                            "ExitNumberUP2='1' " & _
                            "From esporder o " & _
                            "INNER Join orgcompany c " & _
                            "ON c.id = o.companyid " & _
                            "INNER Join orgaddress a " & _
                            "ON a.id = o.shiptoaddressID " & _
                            "INNER Join ebxproductdesign p " & _
                            "ON p.id = o.productdesignID " & _
                            "INNER Join espunitizingspec s " & _
                            "on o.companyid = c.id and o.jobnumber = s.jobnumber " & _
                            "INNER Join ebxunitizingdata u " & _
                            "ON p.unitizingdataID = u.id " & _
                            "where o.jobnumber='" & _JobNumber & "'"

            Dim objDS As DataSet = comm.FillDatasetSQL(strSQL, CommandType.Text, Nothing)

            Dim dr As DataRow
            'Loop throw every row
            For Each dr In objDS.Tables(0).Rows

                'Was the file Exported before
                If Not Exported Then
                    sb.Append("CHPCS") 'Transaction request
                Else
                    sb.Append("CHPCS") 'Transaction request
                End If
                sb.Append("".ToString.PadRight(5, " ")) 'Transaction answer
                sb.Append(Mid(dr.Item("jobnumber").ToString, 1, 10).PadRight(10, " ")) 'WORKIP.job_number
                sb.Append(Mid(dr.Item("SeriesNumber").ToString, 1, 1).PadLeft(1, "0")) 'WORKIP.series_number
                sb.Append("".PadRight(7, " ")) 'spare
                sb.Append(Mid(dr.Item("StepNumber").ToString, 1, 2).PadRight(2, " ")) 'WORKIP.step_number
                sb.Append(_Machine.PadLeft(4, "0")) 'WORKIP.machine_number
                sb.Append("".PadRight(9, " ")) 'Dummy
                sb.Append(Mid(dr.Item("EntryNumberUP1").ToString, 1, 3).PadRight(3, " ")) 'WORKIP.entry_number_up(1)
                sb.Append(Mid(dr.Item("EntryNumberUP2").ToString, 1, 3).PadRight(3, " ")) 'WORKIP.entry_number_up(2)
                sb.Append("".PadRight(8, " ")) 'Dummy
                sb.Append(Mid(dr.Item("ExitNumberUP1").ToString, 1, 3).PadRight(3, " ")) 'WORKIP.exit_number_up(1)
                sb.Append(Mid(dr.Item("ExitNumberUP2").ToString, 1, 3).PadRight(3, " ")) 'WORKIP.exit_number_up(2)
                sb.Append("".PadRight(255, " ")) 'Dummy
                sb.Append("".PadRight(1, " ")) 'EndFeed
                sb.Append(vbCr) 'End

            Next
            _log.Info("Created ADPCS for Machine: ")
            _log.Info(sb.ToString)
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create the ADPCS records..")
        Finally
            _log.Info("Created the ADPCS records..")

        End Try
    End Sub
#End Region

#Region "Create File"
    Private Sub WriteFile()
        Try
            Dim sw As New StreamWriter(_Filename, False, Encoding.Default)
            sw.Write(_sb)
            sw.Flush()
            sw.Close()
            sw.Dispose()
            _log.Info("Finished creating file for JobNumber: " & _JobNumber)
        Catch ex As Exception
            Throw New ApplicationException(ex.Message, ex.InnerException)
            _log.Error("Cannot Create text file..")
        End Try
    End Sub

#End Region

End Class
