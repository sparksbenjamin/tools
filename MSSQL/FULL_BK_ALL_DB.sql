DECLARE @name VARCHAR(50) -- database name 
DECLARE @dir_root VARCHAR(256) -- path for top root of backup files 
DECLARE @path VARCHAR(256) -- path for backup files 
DECLARE @fileName VARCHAR(256) -- filename for backup 
DECLARE @fileDate VARCHAR(20) -- used for file name 
DECLARE @msg NVARCHAR(MAX) -- used for internal messaging and logging

/*
****************************************************
****************************************************
**   MAKE SURE TO SET THE @dir_root and access    **
**                                                **
****************************************************
****************************************************
*/

SET @dir_root = 'F:\Backups\'





SELECT @fileDate = CONVERT(VARCHAR(20),GETDATE(),112) 

DECLARE db_cursor CURSOR FOR 
SELECT name 
FROM MASTER.dbo.sysdatabases 
WHERE name NOT IN ('master','model','msdb','tempdb') 

OPEN db_cursor  
FETCH NEXT FROM db_cursor INTO @name  

WHILE @@FETCH_STATUS = 0  
BEGIN  
	BEGIN TRY
		set @path = @dir_root + @@SERVERNAME + '\' + @name + '\' + 'FULL\'
		exec xp_create_subdir @path;
		SET @fileName = @path + @@SERVERNAME + '_' + @name + '_FULL_'  + @fileDate + '.BAK' 
		BACKUP DATABASE @name TO DISK = @fileName WITH COMPRESSION, NAME='Full Nightly'
		RESTORE VERIFYONLY
		FROM DISK = @fileName
	end try
    begin catch
        set @msg = 'something went wrong!!! with backing up DB=: ' + @name + '    ' + error_message()
        raiserror(@msg,0,0);
    end catch
    FETCH NEXT FROM db_cursor INTO @name 
END 

CLOSE db_cursor  
DEALLOCATE db_cursor
