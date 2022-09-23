declare @tbl as varchar(255)
declare @type as varchar(255)
declare @index as varchar(255)
declare @frag_per as varchar(255)
DECLARE @DB as varchar(100)
set @DB = 'JonasNET'


DECLARE db_cursor CURSOR FOR 



SELECT 
    t.NAME 'Table name',
	i.NAME 'Index_name',
	ips.index_type_desc,
	AVG_FRAGMENTATION_IN_PERCENT
FROM 
    sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'DETAILED') ips
INNER JOIN  
    sys.tables t ON ips.OBJECT_ID = t.Object_ID
INNER JOIN  
    sys.indexes i ON ips.index_id = i.index_id AND ips.OBJECT_ID = i.object_id
WHERE
    AVG_FRAGMENTATION_IN_PERCENT > 3
	AND fragment_count > 5
	AND ips.page_count > 1000
--	AND index_type_desc = 'HEAP'
ORDER BY
    AVG_FRAGMENTATION_IN_PERCENT DESC, fragment_count

OPEN db_cursor  
FETCH NEXT FROM db_cursor INTO @tbl,@index,@type,@frag_per

WHILE @@FETCH_STATUS = 0  
BEGIN
	declare @alter varchar(255)
	set @alter = 'ALTER '
	if @type = 'HEAP'
		set @alter =@alter + 'TABLE ' + @tbl + ' REBUILD'
	ELSE IF @frag_per < 30.00
--		set @alter = @alter + ' INDEX ' 
		set @alter = @alter + 'INDEX ' +  @index + ' ON ' + @tbl + ' REORGANIZE'
	ELSE
		set @alter = @alter + 'INDEX ' + @index + ' ON ' + @tbl + ' REBUILD'
	print 'Working on ' + @tbl
	
	print @alter
	BEGIN TRY  
		exec(@alter)
	END TRY  
	BEGIN CATCH 
		print 'Failed State:'
		set @alter = 'USE ' + @DB + '; ALTER INDEX ALL ON ' + @tbl + ' REBUILD'
		print @alter
		exec @alter
	END CATCH
    FETCH NEXT FROM db_cursor INTO @tbl,@index,@type,@frag_per
END 

CLOSE db_cursor  
DEALLOCATE db_cursor
