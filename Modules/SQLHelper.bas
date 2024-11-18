' SQLHelper.bas
Option Explicit

' Builds a metadata query to fetch column names and data types.
' @param tableName - The name of the table to fetch metadata for
' @param owners - Optional, a list of schema owners to filter by
' @return - A SQL query string
Public Function BuildMetadataQuery(tableName As String, Optional owners As Variant = Array("ZZZHSO", "ZZZHSF")) As String
    Dim ownerFilter As String
    Dim i As Long
    
    ' Construct owner filter
    If Not IsArray(owners) Then
        ownerFilter = "'" & owners & "'"
    Else
        For i = LBound(owners) To UBound(owners)
            If i > LBound(owners) Then ownerFilter = ownerFilter & ", "
            ownerFilter = ownerFilter & "'" & owners(i) & "'"
        Next i
    End If

    ' Build the query
    BuildMetadataQuery = "SELECT COLUMN_NAME, DATA_TYPE FROM ALL_TAB_COLUMNS " & _
                         "WHERE OWNER IN (" & ownerFilter & ") " & _
                         "AND TABLE_NAME = '" & UCase(tableName) & "'"
End Function

' Builds a simple SELECT * query for a given table name.
' @param tableName - The name of the table to fetch data from
' @return - A SQL query string
' I may be able to code this so ownerfilter generates a string of owners separated by commas
Public Function BuildSelectQuery(tableName As String) As String
    BuildSelectQuery = "SELECT * FROM " & UCase(tableName)
End Function


' Builds a SELECT query with specific columns and optional filters.
' @param tableName - The name of the table
' @param columns - An array of column names to include
' @param filters - An optional dictionary of filters (column as key, value as filter condition)
' @return - A SQL query string
Public Function BuildSelectQueryWithFilters(tableName As String, columns As Variant, Optional filters As Object) As String
    Dim columnList As String
    Dim i As Long
    Dim filterClause As String
    Dim key As Variant

    ' Construct column list
    For i = LBound(columns) To UBound(columns)
        If i > LBound(columns) Then columnList = columnList & ", "
        columnList = columnList & columns(i)
    Next i

    ' Construct filter clause
    If Not filters Is Nothing Then
        filterClause = " WHERE "
        For Each key In filters.Keys
            If Len(filterClause) > 7 Then filterClause = filterClause & " AND "
            filterClause = filterClause & key & " = '" & filters(key) & "'"
        Next key
    End If

    ' Build the query
    BuildSelectQueryWithFilters = "SELECT " & columnList & " FROM " & UCase(tableName) & filterClause
End Function
