'Angel Nava
'Spring 2025
'RCET2265
'Car Rental
'Link
Option Explicit On
Option Strict On
Option Compare Binary
Imports System.Security.Policy

Public Class RentalForm

    Sub SetDefaults()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""

        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""


        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False

        NameTextBox.Focus()
    End Sub

    Function CustomerSummary(counting As Boolean) As Integer
        Static _customerSummary As Integer

        If counting = True Then
            _customerSummary += 1
        End If

        Return _customerSummary
    End Function

    Function MileSummary(counting As Boolean, miles As Integer) As Integer
        Static _mileSummary As Integer

        If counting = True Then
            _mileSummary += miles
        End If

        Return _mileSummary
    End Function

    Function ChargeSummary(counting As Boolean, charge As Integer) As Integer
        Static _chargeSummary As Integer

        If counting = True Then
            _chargeSummary += charge
        End If

        Return _chargeSummary
    End Function

    Function UserInputIsValid() As Boolean
        Dim valid As Boolean = True
        Dim message As String = ""

        'Number of Days Schtuff

        If DaysTextBox.Text = "" Then
            valid = False
            message &= "The Number of Days is required" & vbNewLine
            DaysTextBox.Focus()
        Else
            If IsNumeric(DaysTextBox.Text) = False Then
                valid = False
                message &= "The Days Input is not a number" & vbNewLine
                DaysTextBox.Focus()
            Else
                If CInt(DaysTextBox.Text) < 0 Then
                    valid = False
                    MsgBox("The Number of Days must be greater than 0")
                    DaysTextBox.Focus()
                End If
                If CInt(DaysTextBox.Text) > 45 Then
                    valid = False
                    MsgBox("The Number of Days cannot be greater than 45")
                    DaysTextBox.Focus()
                End If
            End If
        End If

        'Odometer Schtuff
        'Beginning odometer reading must be less than ending odometer reading

        If EndOdometerTextBox.Text = "" Then
            valid = False
            message &= "The Ending Odometer Reading is required" & vbNewLine
            EndOdometerTextBox.Focus()
        Else
            If IsNumeric(EndOdometerTextBox.Text) = False Then
                valid = False
                message &= "The Ending Odometer Reading Input is not a number" & vbNewLine
                DaysTextBox.Focus()
            End If
        End If

        If BeginOdometerTextBox.Text = "" Then
            valid = False
            message &= "The Beginning Odometer Reading is required" & vbNewLine
            BeginOdometerTextBox.Focus()
        Else
            If IsNumeric(BeginOdometerTextBox.Text) = False Then
                valid = False
                message &= "The The Beginning Odometer Reading Input is not a number" & vbNewLine
                DaysTextBox.Focus()
            Else
                If CInt(BeginOdometerTextBox.Text) > CInt(EndOdometerTextBox.Text) Then
                    valid = False
                    message &= "The Beginning Odometer reading must be less than Ending Odometer Reading" & vbNewLine
                    BeginOdometerTextBox.Focus()
                End If
            End If
        End If

        'Other

        If IsNumeric(DaysTextBox.Text) = False Then
            valid = False
            message &= "The ZipCode Input is not a number" & vbNewLine
            DaysTextBox.Focus()
        Else
            If ZipCodeTextBox.Text = "" Then
                valid = False
                message &= "Your ZipCode is required" & vbNewLine
                ZipCodeTextBox.Focus()
            End If
        End If
        If StateTextBox.Text = "" Then
            valid = False
            message &= "Your State is required" & vbNewLine
            StateTextBox.Focus()
        End If
        If CityTextBox.Text = "" Then
            valid = False
            message &= "Your City is required" & vbNewLine
            CityTextBox.Focus()
        End If
        If AddressTextBox.Text = "" Then
            valid = False
            message &= "Your Address is required" & vbNewLine
            AddressTextBox.Focus()
        End If
        If NameTextBox.Text = "" Then
            valid = False
            message &= "Your Name is required" & vbNewLine
            NameTextBox.Focus()
        End If

        If valid = False Then
            MsgBox(message, MsgBoxStyle.Critical, "Input Error(s)")
            SetDefaults()
        End If

        Return valid
    End Function


    Function MileageCharge() As Double
        Dim _mileageCharge As Double = 0
        Dim totalMiles As Double = CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)
        Dim twelveCentMiles As Double = 0
        Dim tenCentMiles As Double = 0

        'Mile Cost
        If MilesradioButton.Checked = True Then
            Select Case totalMiles
                Case 0 To 200
                    _mileageCharge = 0
                Case 201 To 500
                    twelveCentMiles = totalMiles - 200
                    _mileageCharge += (twelveCentMiles * 0.12)
                Case Else
                    tenCentMiles = totalMiles - 499
                    twelveCentMiles = totalMiles - tenCentMiles - 200
                    _mileageCharge += (twelveCentMiles * 0.12)
                    _mileageCharge += (tenCentMiles * 0.1)
            End Select
        End If

        'Kilometer Cost
        If KilometersradioButton.Checked = True Then
            totalMiles = (totalMiles / 0.62)

            Select Case totalMiles
                Case 0 To 200
                    _mileageCharge = 0
                Case 201 To 500
                    twelveCentMiles = totalMiles - 200
                    _mileageCharge += (twelveCentMiles * 0.12)
                Case Else
                    tenCentMiles = totalMiles - 499
                    twelveCentMiles = totalMiles - tenCentMiles - 200
                    _mileageCharge += (twelveCentMiles * 0.12)
                    _mileageCharge += (tenCentMiles * 0.1)
            End Select
        End If

        Return _mileageCharge
    End Function

    Function DailyCharge() As Double
        Dim cash As Double = 0

        cash = CInt(DaysTextBox.Text) * 15

        Return cash

    End Function

    'Event Handlers-----------------------------------------------------------------------------------------------------------------

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim totalCharge As Double = 0
        Dim discount As Double = 0

        If UserInputIsValid() Then
            'Put events in here

            If MilesradioButton.Checked = True Then
                TotalMilesTextBox.Text = CStr(CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)) + " mi"
                MileSummary(True, CInt(CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)))
            End If
            If KilometersradioButton.Checked = True Then
                TotalMilesTextBox.Text = CStr((CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)) / 0.62) + " Km"
                MileSummary(True, CInt((CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)) / 0.62))
            End If

            MileageChargeTextBox.Text = "$" + CStr(MileageCharge())
            DayChargeTextBox.Text = "$" + CStr(DailyCharge())
            totalCharge = MileageCharge() + DailyCharge()

            'Discounts
            If AAAcheckbox.Checked = True Then
                discount += 0.05
            End If
            If Seniorcheckbox.Checked = True Then
                discount += 0.03
            End If
            TotalDiscountTextBox.Text = "-$" + CStr(totalCharge * discount)
            If discount > 0 Then
                totalCharge = totalCharge - (totalCharge * discount)
            End If
            'Total Charge
            TotalChargeTextBox.Text = "$" + CStr(totalCharge)
            ChargeSummary(True, CInt(totalCharge))

            'summary schtuff
            CustomerSummary(True)

        End If


    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        Dim message As String = ""

        message &= "Number of Customers: " & CStr(CustomerSummary(False)) & vbNewLine
        message &= "Total Miles ran : " & CStr(MileSummary(False, 0)) & " mi" & vbNewLine
        message &= "Total Charge : $" & CStr(ChargeSummary(False, 0)) & vbNewLine
        MsgBox(message)
        SetDefaults()
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim response = MsgBox("Do you wish to close the program?", MsgBoxStyle.YesNo, "Close Program?")

        ' Take some action based on the response.
        If response = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        SetDefaults()
    End Sub
End Class
