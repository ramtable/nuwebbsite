Option Compare Text
Option Explicit On
Imports JSONkit
Imports ScraperKit
Imports System.IO 'for streamreader in MapCRBR
Imports System.Threading

Module Start
    Sub Main()
        Dim args As String() = Environment.GetCommandLineArgs()
        Dim choice
        If args.Length > 1 Then
            choice = args(1)
        Else
            Console.WriteLine("Select a task to run:")
            Console.WriteLine("1. BuyBacks")
            Console.WriteLine("2. CCASS")
            Console.WriteLine("3. CR")
            Console.WriteLine("4. GetFinancialReports")
            Console.WriteLine("5. HKEXdata")
            Console.WriteLine("6. HKflights")
            Console.WriteLine("7. HKlawSoc")
            Console.WriteLine("8. HKMA")
            Console.WriteLine("9. housing")
            Console.WriteLine("10. ImmD")
            Console.WriteLine("11. LandReg")
            Console.WriteLine("12. Listing")
            Console.WriteLine("13. Quarantine")
            Console.WriteLine("14. Quotes")
            Console.WriteLine("15. SDI")
            Console.WriteLine("16. SFC")
            Console.WriteLine("17. Transport")
            Console.WriteLine("18. Treasury")
            Console.WriteLine("19. UKCH")
            Console.Write("Enter number: ")
            choice = Console.ReadLine()
        End If
        Select Case choice
            Case "1" : Buybacks.Main()
            Case "2" : CCASS.Main()
            Case "3" : CR.Main()
            Case "4" : GetFinancialReports.Main()
            Case "5" : HKEXdata.Main()
            Case "6" : HKflights.Main()
            Case "7" : HKlawSoc.Main()
            Case "8" : HKMA.Main()
            Case "9" : housing.Main()
            Case "10" : ImmD.Main()
            Case "11" : LandReg.Main()
            Case "12" : Listing.Main()
            Case "13" : Quarantine.Main()
            Case "14" : Quotes.Main()
            Case "15" : SDI.Main()
            Case "16" : SFC.Main()
            Case "17" : Transport.Main()
            Case "18" : Treasury.Main()
            Case "19" : UKCH.Main()
            Case Else : Console.WriteLine("Invalid choice.")
        End Select
    End Sub
End Module

