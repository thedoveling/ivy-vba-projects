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
    filters.Add "JOB_ID", "IT_PROG"

    ' Build the query
    sqlQuery = SQLHelper.BuildSelectQueryWithFilters(tableName, columns, filters)
    Debug.Print sqlQuery
End Sub
