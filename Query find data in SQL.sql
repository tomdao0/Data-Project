DECLARE @con VARCHAR(50)
SET @con = ''
DECLARE
@schema_name VARCHAR(100)
,@table_name VARCHAR(100)
,@column_name VARCHAR(100)
,@table_namef VARCHAR(500)
,@constr VARCHAR(100)
SELECT @constr = 'like ''%hi%''' ---string or number or date need to find
IF OBJECT_ID('tempdb..#Results') IS NOT NULL
DROP TABLE #Results
CREATE TABLE #Results (
schema_name VARCHAR(100)
,table_name VARCHAR(100)
,column_name VARCHAR(100)
,condition VARCHAR(50)
,total_row INT
,searched_row INT )
DECLARE tom_cursor CURSOR FOR
SELECT
C.Table_Schema
,C.Table_Name
,C.Column_Name
,'[' + C.Table_CataLog + ']' + '.[' + C.Table_Schema + '].' + '[' + C.Table_Name + ']' AS FullQualifiedTableName
FROM information_schema.Columns C
INNER JOIN information_Schema.Tables T ON C.Table_Name = T.Table_Name
AND T.Table_Type = 'BASE TABLE' AND C.Data_Type IN('varchar','nvarchar') --datatype of column need to find
OPEN tom_cursor
FETCH NEXT FROM tom_cursor INTO
@schema_name
,@table_name
,@column_name
,@table_namef
WHILE @@FETCH_STATUS = 0
BEGIN
DECLARE @SQL VARCHAR(MAX) = NULL
SET @SQL =
'Select ''' + @schema_name + ''' AS schema_name,
''' + @table_name + ''' AS table_name,
''' + @column_name + ''' AS column_name,
''' + @con + ''',(Select count(*) from ' + @table_namef + ' with (nolock)) AS total_row,
count(*) as SearchRowCount from ' + @table_namef + ' with (nolock)
Where ' + @column_name + ' ' + @constr
--print(@SQL)
INSERT INTO #Results
EXEC (@SQL)
FETCH NEXT FROM tom_cursor INTO
@schema_name
,@table_name
,@column_name
,@table_namef
END
CLOSE tom_cursor
DEALLOCATE tom_cursor
SELECT * FROM #Results WHERE searched_row <> 0