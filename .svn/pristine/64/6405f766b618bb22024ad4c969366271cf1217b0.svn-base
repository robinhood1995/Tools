'Option Strict Off
'Option Explicit On
Imports log4net
Imports log4net.Config
Imports System.Configuration
Imports System.Reflection

Public Class clsJulianDates

    Private Shared ReadOnly _log As ILog = LogManager.GetLogger(GetType(clsJulianDates))

#Region " CONSTRUCTORS"
    ''' <summary>
    '''  Manual Constructor
    ''' </summary>
    ''' <remarks>All has to be passed manually</remarks>
    Sub New()
        _log.Info("Starting " & MethodBase.GetCurrentMethod().ToString())
    End Sub
#End Region

#Region " DECLARATIONS"
    Private _mlJulianDate As Integer 'in 'long' format
    Private _msJulianDate As String 'in 'string' format
    Private _mdDate As Date

    ' read only properties
    Private _mlYear As Integer
    Private _mlDay As Integer
    Private _mlDaysInYear As Integer
#End Region

#Region " PROPERTIES"
    Public ReadOnly Property lDaysInYear() As Integer
        Get
            lDaysInYear = _mlDaysInYear
        End Get
    End Property

    Public ReadOnly Property lYear() As Integer
        Get
            lYear = _mlYear
        End Get
    End Property

    Public ReadOnly Property lDay() As Integer
        Get
            lDay = _mlDay
        End Get
    End Property
    Public Property sJulianDate() As String
        Get
            sJulianDate = _msJulianDate
        End Get
        Set(ByVal Value As String)
            _msJulianDate = Value
        End Set
    End Property

    Public Property lJulianDate() As Integer
        Get
            lJulianDate = _mlJulianDate
        End Get
        Set(ByVal Value As Integer)
            _mlJulianDate = Value
        End Set
    End Property

    Public Property dDate() As Date
        Get
            dDate = _mdDate
        End Get
        Set(ByVal Value As Date)
            _mdDate = Value
        End Set
    End Property
#End Region

    Public Function JulianToDate(Optional ByRef vntJulianDate As Object = Nothing) As Date
        If Not IsNothing(vntJulianDate) Then 'did they supply the Julian date in the function
            lJulianDate = vntJulianDate
        End If
        If Len(CStr(lJulianDate)) > 5 Then
            Err.Raise(vbObjectError + 2, "clsJulianToDate:JulianToDate", "Julian date greater than 5 characters.")
            Exit Function
        ElseIf Len(CStr(lJulianDate)) < 1 Then
            Err.Raise(vbObjectError + 3, "clsJulianToDate:JulianToDate", "Julian date less than one characters.")
            Exit Function
        End If

        _mlYear = lJulianDate \ 1000 'get the year part
        _mlDay = lJulianDate - lYear * 1000 'get the day of the year part
        _mlDaysInYear = DaysInTheYear(DateSerial(lYear, 1, 1)) 'number of days in the year
        If _mlDay >= 1 And _mlDay <= lDaysInYear Then 'within the range?
            _mdDate = System.DateTime.FromOADate(DateSerial(lYear, 1, 1).ToOADate + lDay - 1) 'yes, return what we found
            JulianToDate = _mdDate 'and return in the function
        Else
            Err.Raise(vbObjectError + 1, "clsJulianToDate:JulianToDate", "Invalid Julian day, less than 1 or greater than " & lDaysInYear & ".")
        End If

        Return JulianToDate
    End Function

    Public Function DateToJulian(Optional ByRef vntDate As Object = Nothing) As String
        If Not IsNothing(vntDate) Then 'did they supply the date in the function?
            dDate = vntDate
        End If
        _mlYear = Year(dDate) 'get the year part
        _mlDay = DateDiff(Microsoft.VisualBasic.DateInterval.DayOfYear, DateSerial(lYear, 1, 1), dDate) + 1 'day part
        lJulianDate = CInt(Right(Format(lYear, "0000"), 2) & Format(lDay, "000")) 'convert to yyddd
        sJulianDate = Format(lJulianDate, "00000") 'format as string "yyddd"
        DateToJulian = sJulianDate 'return the string
    End Function

    Public Function DaysInTheYear(ByRef dDate As Date) As Integer
        ' Return the number of days in the year using dDate
        DaysInTheYear = DateDiff(Microsoft.VisualBasic.DateInterval.DayOfYear, DateSerial(Year(dDate), 1, 1), LastDayOfTheYear(dDate)) + 1
    End Function

    Public Function LastDayOfTheYear(ByRef dDate As Date) As Date
        ' Return the last day of the year using dDate
        LastDayOfTheYear = System.DateTime.FromOADate(DateSerial(Year(dDate) + 1, 1, 1).ToOADate - 1) 'one day less than January 1 the next year
    End Function
End Class