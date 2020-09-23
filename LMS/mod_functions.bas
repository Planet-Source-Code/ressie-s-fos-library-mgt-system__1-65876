Attribute VB_Name = "mod_functions"
Function loadname()
    frmSearchName.lstSearchResult.ListItems.Clear
    frmSearchName.adoqry_Clients.Refresh
    Do While Not frmSearchName.adoqry_Clients.Recordset.EOF
    Dim m As Integer
    m = frmSearchName.lstSearchResult.ListItems.Count + 1
    Set lst = frmSearchName.lstSearchResult.ListItems.Add(, , m & ".  " & frmSearchName.adoqry_Clients.Recordset.Fields!Fullname, , 0)
        lst.SubItems(1) = frmSearchName.adoqry_Clients.Recordset.Fields!RefNo
        frmSearchName.adoqry_Clients.Recordset.MoveNext
    Loop
    frmSearchName.adoqry_Clients.Refresh
End Function
Function listbooks()
    
    frmBorrow.adoRefNo.Refresh
    frmBorrow.adoRefNo.Recordset.Fields!RefNo = frmBorrow.lblRefNo
    frmBorrow.adoRefNo.Recordset.Update
    frmBorrow.adoRefNo.Refresh
    
    '=========================
    frmBorrow.lstBooksBorrowed.ListItems.Clear
    frmBorrow.adoqryListBooks.Refresh
    Do While Not frmBorrow.adoqryListBooks.Recordset.EOF
    Set lst = frmBorrow.lstBooksBorrowed.ListItems.Add(, , frmBorrow.adoqryListBooks.Recordset.Fields!BookCode, , 0)
        lst.SubItems(1) = frmBorrow.adoqryListBooks.Recordset.Fields!Title
        lst.SubItems(2) = frmBorrow.adoqryListBooks.Recordset.Fields!Author
        lst.SubItems(3) = frmBorrow.adoqryListBooks.Recordset.Fields!DateBorrowed
        lst.SubItems(4) = frmBorrow.adoqryListBooks.Recordset.Fields!NoCopyBorrowed
        lst.SubItems(5) = frmBorrow.adoqryListBooks.Recordset.Fields!Status
        frmBorrow.adoqryListBooks.Recordset.MoveNext
    Loop
    frmBorrow.adoqryListBooks.Refresh

End Function
