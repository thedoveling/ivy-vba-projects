Sub TestFilteredQuery()
    Dim tableName As String
    Dim columns As Variant
    Dim filters As Object
    Dim sqlQuery As String

    ' Table and columns
    tableName = "EMPLOYEES"
    columns = Array("EMPLOYEE_ID", "FIRST_NAME", "DEPARTMENT_ID")

    ' Filters
    Set filters = CreateObject("Scripting.Dictionary")
    filters.Add "DEPARTMENT_ID", "10"
    filters.Add "CATEGORY", "ADDRESS_QUERIES"

    ' Build the query
    sqlQuery = SQLHelper.BuildSelectQueryWithFilters(tableName, columns, filters, "schema_name")
    Debug.Print sqlQuery
End Sub


' Here you don't need to worry about schema name, just add these as prefixes to the table names in the joins dictionary and query string.
Sub TestJoinQuery()
    Dim joins As Object
    Dim filters As Object
    Dim columns As Variant
    Dim query As String

    ' Define the columns to select
    columns = Array("a.COLUMN_1", "b.COLUMN_2", "c.COLUMN_3")

    ' Define the joins
    Set joins = CreateObject("Scripting.Dictionary")
    joins.Add "TableB b", "a.ID = b.ID"
    joins.Add "TableC c", "b.OtherID = c.OtherID"

    ' Define the filters
    Set filters = CreateObject("Scripting.Dictionary")
    filters.Add "a.Status", "Active"

    ' Build the query
    query = BuildJoinQuery("TableA a", joins, columns, filters)
End Sub