' SQLHelper.bas
Option Explicit

' Builds a metadata query to fetch column names and data types.
' @param tableName - The name of the table to fetch metadata for
' @param owners - Optional, a comma-separated string of schema owners (default: "ZZZHSO")
' @return - A SQL query string
Option Explicit
Option Explicit

' Builds a metadata query to fetch column names and data types.
Public Function BuildMetadataQuery(tableName As String, Optional owners As String = "ZZZHSO") As String
    Dim ownerFilter As String
    Dim ownerArray() As String
    Dim i As Long

    ownerArray = Split(owners, ",")

    For i = LBound(ownerArray) To UBound(ownerArray)
        If i > LBound(ownerArray) Then ownerFilter = ownerFilter & ", "
        ownerFilter = ownerFilter & "'" & Trim(ownerArray(i)) & "'"
    Next i

    BuildMetadataQuery = "SELECT COLUMN_NAME, DATA_TYPE FROM ALL_TAB_COLUMNS " & _
                         "WHERE OWNER IN (" & ownerFilter & ") " & _
                         "AND TABLE_NAME = '" & UCase(tableName) & "'"
End Function

' Builds a SELECT * query for a table.
Public Function BuildSelectQuery(tableName As String) As String
    BuildSelectQuery = "SELECT * FROM " & UCase(tableName)
End Function

' Builds a SELECT query with filters.
Public Function BuildSelectQueryWithFilters(tableName As String, columns As Variant, Optional filters As Object, Optional schema As String = "") As String
    Dim columnList As String, filterClause As String, fullTableName As String
    Dim i As Long, key As Variant

    If schema <> "" Then fullTableName = schema & "." & tableName Else fullTableName = tableName

    For i = LBound(columns) To UBound(columns)
        If i > LBound(columns) Then columnList = columnList & ", "
        columnList = columnList & columns(i)
    Next i

    If Not filters Is Nothing Then
        filterClause = " WHERE "
        For Each key In filters.Keys
            If Len(filterClause) > 7 Then filterClause = filterClause & " AND "
            filterClause = filterClause & key & " = '" & filters(key) & "'"
        Next key
    End If

    BuildSelectQueryWithFilters = "SELECT " & columnList & " FROM " & fullTableName & filterClause
End Function


' Builds a SELECT query with JOINs and optional filters.
' @param baseTable - The main table for the query
' @param joins - A dictionary with keys as table names and values as ON conditions
' @param columns - An array of column names to select
' @param filters - An optional dictionary of filters (column as key, value as filter condition)
' @return - A SQL query string
Public Function BuildJoinQuery(baseTable As String, joins As Object, columns As Variant, Optional filters As Object) As String
    Dim columnList As String
    Dim joinClause As String
    Dim filterClause As String
    Dim i As Long
    Dim key As Variant

    ' Construct column list
    For i = LBound(columns) To UBound(columns)
        If i > LBound(columns) Then columnList = columnList & ", "
        columnList = columnList & columns(i)
    Next i

    ' Construct join clause
    For Each key In joins.Keys
        joinClause = joinClause & " JOIN " & key & " ON " & joins(key)
    Next key

    ' Construct filter clause
    If Not filters Is Nothing Then
        filterClause = " WHERE "
        For Each key In filters.Keys
            If Len(filterClause) > 7 Then filterClause = filterClause & " AND "
            filterClause = filterClause & key & " = '" & filters(key) & "'"
        Next key
    End If

    ' Build the query
    BuildJoinQuery = "SELECT " & columnList & " FROM " & baseTable & joinClause & filterClause
End Function
