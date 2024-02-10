Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports Ikanobanken.BankService.ServiceTypes.Response
Imports Ikanobanken.BankService.ServiceTypes.SubTypes
Imports Ikanobanken.Infrastructure.Bank.Configuration
Imports Ikanobanken.Library.Text
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class iTextSharpHelper

    Public Function CreateAnnualStatementPdfResponse(statement As GetAnnualStatementResponse, socialSecurityNumber As String) As GetAnnualStatementPdfResponse

        If statement Is Nothing OrElse statement.AnnualStatement Is Nothing OrElse statement.AnnualStatement.History.Count = 0 Then
            Throw New ApplicationException("CreateAnnualStatementPdfResponse: No Annual Statement available!")
        End If

        Dim output As Byte()
        Using ms As New MemoryStream()

            Dim doc As New Document(PageSize.A4)
            Dim pWriter As PdfWriter = PdfWriter.GetInstance(doc, ms)
            pWriter.PdfVersion = "3"c

            Dim title As String = "Årsbesked " + statement.AnnualStatement.Balance.YearConcerned
            doc.AddAuthor("Ikano Bank")
            doc.AddSubject(title)
            doc.AddTitle(title)
            doc.AddKeywords(title)
            doc.Open()

            Dim imgUri As String = ConfigurationManager.Instance.AnnualStatementPdfLogoUrl
            'Old "Viggen" url
            '"https://secure.ikanobank.se/templates/_common/branding_secure/img/logos/ikano-bank-print.gif"
            'PROD
            'https://ikanobank.se/-/media/sweden/digitalstore/logotypes/ikano-bank.png
            'UAT
            'https://view9.uat.ikanobank-internal.se/-/media/sweden/digitalstore/logotypes/ikano-bank.png 
            'DEV/TEST
            'https://view9.sit.ikanobank-internal.se/-/media/sweden/digitalstore/logotypes/ikano-bank.png 

            Dim fontHeading As New Font(Font.FontFamily.UNDEFINED, 12, Font.BOLD)
            Dim fontBody As New Font(Font.FontFamily.UNDEFINED, 10, Font.NORMAL)
            Dim fontBodyImportant As New Font(Font.FontFamily.UNDEFINED, 10, Font.BOLD Or Font.ITALIC)

            Dim fpNoDecimals As New System.Globalization.NumberFormatInfo() With {.NumberDecimalDigits = 0}
            Dim fpTwoDecimals As New System.Globalization.NumberFormatInfo With {.NumberDecimalDigits = 2}

            CreatePdfHeader(doc, statement, socialSecurityNumber, imgUri)
            CreatePdfBody(doc, statement, fontHeading, fontBody, fontBodyImportant, fpNoDecimals, fpTwoDecimals)
            CreatePdfFooter(doc, statement, fontHeading, fontBody, fontBodyImportant, fpNoDecimals, fpTwoDecimals)

            doc.Close()

            output = ms.ToArray()

        End Using

        Dim returnValue As New GetAnnualStatementPdfResponse() With {.AnnualStatement = New AnnualStatementResultPdf()}
        returnValue.AnnualStatement.PdfOutput = Convert.ToBase64String(output)

        Return returnValue

    End Function

    Protected Function GetCustomerFormattedAddress(socialSecurityNumber As String) As String

        Dim firstName As String = String.Empty
        Dim lastName As String = String.Empty
        Dim careOfAddress As String = String.Empty
        Dim streetAddress As String = String.Empty
        Dim postalCode As String = String.Empty
        Dim postalAddress As String = String.Empty

        Dim customer As Ikanobanken.Infrastructure.Common.DataSets.Customer = Nothing
        Using customerServices As New Ikanobanken.Infrastructure.Common.EntityServices.CustomerServices()
            customer = customerServices.GetCustomer(socialSecurityNumber)
        End Using
        If customer Is Nothing Then
            Throw New ApplicationException("GetCustomerFormattedAddress: No Customer data available!")
        End If

        If Not customer.Customers(0).IsFirstNameNull() Then
            firstName = Format.FormatCustomerName(customer.Customers(0).FirstName)
        End If
        lastName = Format.FormatCustomerName(customer.Customers(0).SurName)
        streetAddress = Format.FormatCustomerName(customer.Customers(0).Address)
        If Not customer.Customers(0).IsCareOfAddressNull() Then
            careOfAddress = Format.FormatCustomerName(customer.Customers(0).CareOfAddress)
        End If
        If Not customer.Customers(0).IsPostalCodeNull() Then
            postalCode = Format.FormatCustomerName(customer.Customers(0).PostalCode.ToString())
        End If
        If Not customer.Customers(0).IsPostalAddressNull() Then
            postalAddress = Format.FormatCustomerName(customer.Customers(0).PostalAddress)
        End If

        If Not careOfAddress = String.Empty Then
            careOfAddress += Environment.NewLine
        End If

        Dim address As String = String.Format("{1}{0}{2}{3}{0}{4}{0}", Environment.NewLine,
                                              firstName + " " + lastName,
                                              streetAddress,
                                              careOfAddress,
                                              postalCode + " " + postalAddress)
        Return address

    End Function

    Protected Sub CreatePdfHeader(ByRef doc As Document,
                                  getAnnualStatementResponse1 As GetAnnualStatementResponse,
                                  socialSecurityNumber As String,
                                  imgUri As String)

        Dim tbl As PdfPTable = New PdfPTable(2)
        tbl.SpacingAfter = 20
        tbl.DefaultCell.Border = PdfPCell.NO_BORDER
        tbl.WidthPercentage = 100
        tbl.SetWidths(New Single() {1, 1})

        Dim gif As Image = Image.GetInstance(New Uri(imgUri))
        gif.ScalePercent(38)

        Dim imgCell As New PdfPCell(gif) With {.Border = PdfPCell.NO_BORDER}
        imgCell.PaddingTop = 6
        tbl.AddCell(imgCell)

        Dim CustomerAdressInfo As New Paragraph(GetCustomerFormattedAddress(socialSecurityNumber))
        CustomerAdressInfo.Alignment = 2

        Dim CustomerAdressCell As New PdfPCell() With {.Border = PdfPCell.NO_BORDER}

        CustomerAdressCell.AddElement(CustomerAdressInfo)

        tbl.AddCell(CustomerAdressCell)



        Dim font As New Font(Font.FontFamily.UNDEFINED, 16, Font.BOLD)
        Dim cell As New PdfPCell(New Phrase("Årsbesked " +
                    getAnnualStatementResponse1.AnnualStatement.Balance.YearConcerned, font)) With {.Border = PdfPCell.NO_BORDER}

        tbl.AddCell(cell)
        tbl.AddCell(New Phrase(String.Empty))

        Dim fontBody As New Font(font.FontFamily.UNDEFINED, 10, font.NORMAL)
        AddTableRow(String.Empty, tbl, fontBody)
        AddTableRow(String.Empty, tbl, fontBody)
        AddTableRow("Kontrolluppgifter markerade i fetstil har lämnats till Skatteverket.", tbl, fontBody)
        AddTableRow(String.Empty, tbl, fontBody)
        AddTableRow("Alla inlåningskonton hos Ikano Bank omfattas av den statliga insättningsgarantin.", tbl, fontBody)
        AddTableRow(String.Empty, tbl, fontBody)

        doc.Add(tbl)

    End Sub

    Protected Sub AddTableRows(ByRef tbl As PdfPTable,
                               headingCol1 As String,
                               headingCol2 As String,
                               bodyCol1 As String,
                               bodyCol2 As String,
                               fontHeader As Font,
                               bodyFont As Font)

        Dim cellHeading1 As New PdfPCell(New Phrase(headingCol1, fontHeader)) With {.Border = PdfPCell.NO_BORDER}
        Dim cellHeading2 As New PdfPCell(New Phrase(headingCol2, fontHeader)) With {.Border = PdfPCell.NO_BORDER}

        Dim cellBody1 As New PdfPCell(New Phrase(bodyCol1, bodyFont)) With {.Border = PdfPCell.NO_BORDER}
        Dim cellBody2 As New PdfPCell(New Phrase(bodyCol2, bodyFont)) With {.Border = PdfPCell.NO_BORDER}

        tbl.AddCell(cellHeading1)
        tbl.AddCell(cellHeading2)
        tbl.AddCell(cellBody1)
        tbl.AddCell(cellBody2)

    End Sub

    Protected Sub AddTableRow(col1 As String,
                              ByRef tbl As PdfPTable,
                              font As Font,
                              Optional importantRow As Boolean = False)

        AddTableRow(col1, String.Empty, tbl, font, importantRow)

    End Sub

    Protected Sub AddTableRow(col1 As String,
                              col2 As String,
                              ByRef tbl As PdfPTable,
                              font As Font,
                              Optional importantRow As Boolean = False)

        Dim backgroundColor As BaseColor = BaseColor.WHITE
        If (importantRow) Then
            backgroundColor = New BaseColor(235, 234, 233)
        End If

        AddTableRow(col1, col2, tbl, font, backgroundColor)

    End Sub

    Protected Sub AddTableRow(col1 As String,
                              col2 As String,
                              ByRef tbl As PdfPTable,
                              font As Font,
                              backgroundColor As BaseColor)

        Dim cell1 As New PdfPCell(New Phrase(col1, font)) With {.Border = PdfPCell.NO_BORDER, .BackgroundColor = backgroundColor}

        If (String.IsNullOrEmpty(col2)) Then

            cell1.Colspan = 2
            tbl.AddCell(cell1)

        Else

            Dim cell2 As New PdfPCell(New Phrase(col2, font)) With {.Border = PdfPCell.NO_BORDER, .BackgroundColor = backgroundColor,
                                                                    .HorizontalAlignment = PdfPCell.ALIGN_RIGHT}
            tbl.AddCell(cell1)
            tbl.AddCell(cell2)

        End If


    End Sub

    Protected Function IsAccountClosed(account As History) As Boolean

        Return (account.StatusText.Trim.ToUpper() = "AVSLUTAD") OrElse _
               (account.StatusText.Trim.ToUpper() = "AVSLUTAT")

    End Function

    Protected Sub CheckAndAddAccountClosed(ByRef tbl As PdfPTable, account As History, fontBody As Font)

        If IsAccountClosed(account) Then
            AddTableRow("Kontot avslutat", tbl, fontBody)
        End If

    End Sub

    Protected Sub CreatePdfBody(ByRef doc As Document, getAnnualStatementResponse1 As GetAnnualStatementResponse,
                                fontHeading As Font, fontBody As Font, fontBodyImportant As Font,
                                fpNoDecimals As NumberFormatInfo, fpTwoDecimals As NumberFormatInfo)



        Dim accounts As IList(Of History) = getAnnualStatementResponse1.AnnualStatement.History
        Dim tbl As PdfPTable

        For Each currentAccount As History In accounts.Where(Function(x) x.EngagementType = "D" OrElse x.EngagementType = "T")

            Dim productText As String = currentAccount.ProductText.ToUpper()
            tbl = CreateTable()
            AddTableRow(currentAccount.ProductText + " " + currentAccount.EngagementNumber, tbl, fontHeading)

            If currentAccount.EngagementType = "D" Then

                AddSparkontoAccount(tbl, currentAccount, fontBody, fontBodyImportant, fpNoDecimals, fpTwoDecimals)

            ElseIf productText.IndexOf("TILLVÄXT") > -1 Then

                AddTillvaxtAccount(tbl, currentAccount, fontBody, fontBodyImportant, fpNoDecimals, fpTwoDecimals)

            ElseIf productText.IndexOf("BÖRS") > -1 Then

                AddBorsAccount(tbl, currentAccount, fontBody, fontBodyImportant, fpNoDecimals, fpTwoDecimals)

            ElseIf (productText.IndexOf("SPARKONTO FIX") <> -1) OrElse _
                   (productText.IndexOf("SPARKTO FIX") <> -1) OrElse _
                   (productText.IndexOf("FÖRETAGSKTO FIX") <> -1) Then

                AddSparkontoAccount(tbl, currentAccount, fontBody, fontBodyImportant, fpNoDecimals, fpTwoDecimals)

            End If

            doc.Add(tbl)

        Next

        For Each currentAccount As History In accounts.Where(Function(x) x.EngagementType = "L")

            tbl = CreateTable()
            AddLoanAccount(tbl, currentAccount, fontBody, fontHeading, fontBodyImportant, fpNoDecimals, fpTwoDecimals)
            doc.Add(tbl)

        Next

    End Sub

    Protected Sub AddLoanAccount(tbl As PdfPTable, account As History, fontBody As Font, fontHeading As Font, fontBodyImportant As Font,
                                 fpNoDecimals As NumberFormatInfo, fpTwoDecimals As NumberFormatInfo)

        AddTableRow(account.ProductText + " " + account.EngagementNumber, tbl, fontHeading)
        AddTableRow("Utgiftsränta", account.PaidInterest.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        AddTableRow("Skuld", account.Debth.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        AddTableRow("Din andel av kapital:" + account.CapitalSharePercentage.Value.ToString("N", fpNoDecimals) + "%",
                    account.DebthShare.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        AddTableRow("Din andel av ränta:" + account.PaidInterestSharePercentage.Value.ToString("N", fpNoDecimals) + "%",
                    account.PaidInterestShare.Value.ToString("N", fpTwoDecimals), tbl, fontBodyImportant, True)

        CheckAndAddAccountClosed(tbl, account, fontBody)

    End Sub

    Protected Sub AddSparkontoAccount(tbl As PdfPTable, account As History, fontBody As Font, fontBodyImportant As Font,
                                      fpNoDecimals As NumberFormatInfo, fpTwoDecimals As NumberFormatInfo)

        AddTableRow("Aktuell räntesats", account.Interest1.Value.ToString("N", fpTwoDecimals) + "%", tbl, fontBody)
        AddTableRow("Inkomstränta", account.ReceivedInterest.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        AddTableRow("Tillgodohavande", account.Balance.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        AddTableRow("Preliminärskatt", account.PreliminaryTax.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        AddTableRow("Din andel av kapital: " + account.CapitalSharePercentage.Value.ToString("N", fpNoDecimals) + "%",
                    account.BalanceShare.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        AddTableRow("Din andel av ränta: " + account.ReceivedInterestSharePercentage.Value.ToString("N", fpNoDecimals) + "%",
                    account.ReceivedInterestShare.Value.ToString("N", fpTwoDecimals), tbl, fontBodyImportant, True)

        CheckAndAddAccountClosed(tbl, account, fontBody)

    End Sub

    Protected Sub AddBorsAccount(tbl As PdfPTable, account As History, fontBody As Font, fontBodyImportant As Font,
                                 fpNoDecimals As NumberFormatInfo, fpTwoDecimals As NumberFormatInfo)

        If IsAccountClosed(account) Then
            AddTableRow("Utbetalt belopp", account.InterestOverDesk.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
            AddTableRow("Kapitalvinst", account.ReceivedInterest.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        Else
            AddTableRow("Tillgodohavande", account.Balance.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
        End If

    End Sub

    Protected Sub AddTillvaxtAccount(tbl As PdfPTable, account As History, fontBody As Font, fontBodyImportant As Font,
                                     fpNoDecimals As NumberFormatInfo, fpTwoDecimals As NumberFormatInfo)

        If IsAccountClosed(account) OrElse
            account.ReceivedInterest > 0 OrElse
            account.PreliminaryTax > 0 Then

            AddTableRow("Inkomstränta", account.ReceivedInterest.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
            AddTableRow("Tillgodohavande", account.Balance.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
            AddTableRow("Preliminärskatt", account.PreliminaryTax.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
            AddTableRow("Din andel av kapital: " + account.CapitalSharePercentage.Value.ToString("N", fpNoDecimals) + "%",
                        account.BalanceShare.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
            AddTableRow("Din andel av ränta: " + account.ReceivedInterestSharePercentage.Value.ToString("N", fpNoDecimals) + "%",
                        account.ReceivedInterestShare.Value.ToString("N", fpTwoDecimals), tbl, fontBodyImportant, True)

            CheckAndAddAccountClosed(tbl, account, fontBody)

        Else

            AddTableRow("Tillgodohavande",
                        account.Balance.Value.ToString("N", fpTwoDecimals), tbl, fontBody)
            AddTableRow("Din andel av kapital: " + account.CapitalSharePercentage.Value.ToString("N", fpNoDecimals) + "%",
                        account.BalanceShare.Value.ToString("N", fpTwoDecimals), tbl, fontBodyImportant, True)

        End If

    End Sub

    Protected Sub CreatePdfFooter(ByRef doc As Document, getAnnualStatementResponse1 As GetAnnualStatementResponse,
                                  fontHeading As Font, fontBody As Font, fontBodyImportant As Font,
                                  fpNoDecimals As NumberFormatInfo, fpTwoDecimals As NumberFormatInfo)

        Dim summary As Balance = getAnnualStatementResponse1.AnnualStatement.Balance
        Dim tbl As PdfPTable

        tbl = CreateTable()
        tbl.SpacingBefore = 14

        AddTableRow("Summa inkomstränta", summary.SumReceivedInterest.Value.ToString("N", fpTwoDecimals), tbl, fontBodyImportant, True)
        AddTableRow("Summa innehållen preliminärskatt", summary.SumPreliminaryTax.Value.ToString("N", fpTwoDecimals), tbl, fontBodyImportant, True)
        AddTableRow("Summa utgiftsränta", summary.SumPaidInterest.Value.ToString("N", fpTwoDecimals), tbl, fontBodyImportant, True)
        AddTableRow(String.Empty, tbl, fontBody)

        Dim linkFont As New Font(Font.FontFamily.UNDEFINED, 10, Font.BOLD, New BaseColor(61, 59, 57))

        Dim c As New Chunk("Förklaring till ditt årsbesked", linkFont)
        c.SetUnderline(0.1F, -2.0F)
        c.SetAnchor("https://ikanobank.se/minasidor/arsbesked")

        Dim cell1 As New PdfPCell() With {.Border = PdfPCell.NO_BORDER, .Colspan = 2}
        cell1.AddElement(c)

        tbl.AddCell(cell1)

        doc.Add(tbl)

    End Sub

    Protected Function CreateTable() As PdfPTable

        Dim tbl As PdfPTable = New PdfPTable(2)
        tbl.DefaultCell.Border = PdfPCell.NO_BORDER
        tbl.WidthPercentage = 100
        tbl.SetWidths(New Single() {5, 2})
        tbl.KeepTogether = True
        tbl.SpacingAfter = 10

        Return tbl

    End Function

End Class
