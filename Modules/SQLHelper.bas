' SQLHelper.bas
Option Explicit

' Builds a metadata query to fetch column names and data types.
' @param tableName - The name of the table to fetch metadata for
' @param owners - Optional, a comma-separated string of schema owners (default: "ZZZHSO")
' @return - A SQL query string
Public Function BuildMetadataQuery(tableName As String, Optional owners As String = "ZZZHSO") As String
    Dim ownerFilter As String
    Dim ownerArray() As String
    Dim i As Long

    ' Split owners into an array if necessary
    ownerArray = Split(owners, ",")

    ' Construct owner filter
    For i = LBound(ownerArray) To UBound(ownerArray)
        If i > LBound(ownerArray) Then ownerFilter = ownerFilter & ", "
        ownerFilter = ownerFilter & "'" & Trim(ownerArray(i)) & "'"
    Next i

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
Public Function BuildSelectQueryWithFilters(tableName As String, columns As Variant, Optional filters As Object, Optional schema As String = "") As String
    Dim columnList As String
    Dim i As Long
    Dim filterClause As String
    Dim key As Variant
    Dim fullTableName As String
    
    ' Construct full table name with schema prefix
    If schema <> "" Then
        fullTableName = schema & "." & tableName
    Else
        fullTableName = tableName
    End If
    
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
