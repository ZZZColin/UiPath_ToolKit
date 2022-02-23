Attribute VB_Name = "Module1"
Sub main()
    
    SourceFilePath = ""
    
    FilePath = "C:\Users\colin_zhao\Desktop\AHDCC\NY PKG Mod letter approved by QBE.docx"
    
    Set wapp = CreateObject("word.application")

    Set eapp = CreateObject("excel.application")
    
    Set efile = eapp.Workbooks.Open(SourceFilePath)
    
    Data = efile.Sheets("").UsedRange.Value
    
    eapp.Application.DisplayAlerts = False
    efile.Close
    Set efile = Nothing
    eapp.Application.DisplayAlerts = True
    
    eapp.Quit
    Set eapp = Nothing
    
    For i = 4 To UBound(Data, 1)

        NI = Data(i, 18)
        
        NIAddress = Data(i, 19)
        
        NICity = Data(i, 21)
        
        NIState = Data(i, 22)
        
        NIZip = Data(i, 23)
        
        Agency = Data(i, 24)
        
        AgencyAddress = Data(i, 25)
        
        AgencyCity = Data(i, 27)
        
        AgencyState = Data(i, 28)
        
        AgencyZip = Data(i, 29)
        
        PolicyNumber = Data(i, 1)
        
        PolicyStartDate = Data(i, 13)
        
        PolicyEndDate = Data(i, 15)
        
        Premium = Data(i, 7)
        
        
        Set wfile = wapp.Documents.Open(FilePath)
    
        Content = wfile.Content
        
        'Debug.Print Content
        
        Content = Replace(Content, "*AGENCY*", Agency)
        
        Content = Replace(Content, "*AGENCY ADDRESS*", AgencyAddress)
        
        Content = Replace(Content, "*AGENCY CITY, STATE  ZIP*", AgencyCity & ", " & AgencyState & " " & AgencyZip)
        
        Content = Replace(Content, "*NAMED INSURED*", NI)
        
        Content = Replace(Content, "*INSURED ADDRESS*", NIAddress)
        
        Content = Replace(Content, "*INSURED CITY, STATE ZIP*", NICity & ", " & NIState & " " & NIZip)
        
        Content = Replace(Content, "*Policy number:     *", "Policy number:     " & PolicyNumber)
        
        Content = Replace(Content, "*Policy Period: *", "Policy Period: " & PolicyStartDate & " - " & PolicyEndDate)
        
        Content = Replace(Content, "*Premium Refund:    *", "Premium Refund:    " & Premium)


        wfile.Content = Content
        
        wfile.ExportAsFixedFormat OutputFileName:= _
                                  wfile.Path & "\" & PolicyNumber & ".pdf", _
                                  ExportFormat:=wdExportFormatPDF, _
                                  OpenAfterExport:=False, _
                                  OptimizeFor:=wdExportOptimizeForPrint, _
                                  Range:=wdExportAllDocument, _
                                  IncludeDocProps:=True, _
                                  CreateBookmarks:=wdExportCreateWordBookmarks, _
                                  BitmapMissingFonts:=True
    
        wapp.Application.DisplayAlerts = False
        wfile.Close
        Set wfile = Nothing
        wapp.Application.DisplayAlerts = True
        
    Next i
    
    
    
    
    
    wapp.Quit
    Set wapp = Nothing
    

End Sub
