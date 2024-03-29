USE [AVERYLA]
GO
/****** Object:  StoredProcedure [dbo].[GetAttendanceReport]    Script Date: 2024-03-03 12:06:08 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author		:	Easin Alif
-- Create date:  2024-02-28
-- Description	:	Datewise Attendance Status
-- Notes			: Columns sequence shound not be change, otherwise excel output data will be mismatch
-- =============================================

-- EXEC GetAttendanceReport '2023-02-15',  '2023-02-16'
ALTER PROCEDURE [dbo].[GetAttendanceReport]
    @startDate DATE,
    @endDate DATE
AS
BEGIN
    DECLARE @SQL VARCHAR(MAX) = '', 
            @dateRange VARCHAR(MAX) = '';

    SELECT @dateRange += ',' + QUOTENAME(CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120))
    FROM master..spt_values 
    WHERE type = 'P' AND DATEADD(DAY, number, @startDate) <= @endDate;

    SELECT @dateRange = STUFF(@dateRange, 1, 1, '');

    SET @SQL = '
    SELECT 
        EmployeeId, 
        e.FirstName +'' ''+ e.LastName AS [Name], 
        b.[Name] AS [Location], 
		ISNULL(d.Name,'''') AS [DepartmentName], 
        ISNULL( ds.Name,'''') AS [DesignationName], 
        ' + STUFF((
            SELECT ', CONVERT(VARCHAR, MAX(CASE WHEN AttendanceDate = ''' + CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + ''' THEN CONVERT(VARCHAR(5), StartTime, 108) END)) AS ' + QUOTENAME(CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + 'StartTime') +
		', CONVERT(VARCHAR, MAX(CASE WHEN AttendanceDate = ''' + CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + ''' THEN CONVERT(VARCHAR(5), EndTime, 108) END)) AS ' + QUOTENAME(CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + 'EndTime') +
		', CONVERT(VARCHAR, MAX(CASE WHEN AttendanceDate = ''' + CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + ''' THEN CONVERT(VARCHAR(5), Early, 108) END)) AS ' + QUOTENAME(CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + 'Early') +
		', CONVERT(VARCHAR, MAX(CASE WHEN AttendanceDate = ''' + CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + ''' THEN CONVERT(VARCHAR(5), Late, 108) END)) AS ' + QUOTENAME(CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + 'Late') +
                   ', MAX(CASE WHEN AttendanceDate = ''' + CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + ''' THEN Status END) AS ' + QUOTENAME(CONVERT(VARCHAR(10), DATEADD(DAY, number, @startDate), 120) + 'Status')
            FROM master..spt_values 
            WHERE type = 'P' AND DATEADD(DAY, number, @startDate) <= @endDate
            FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '') + '
    FROM 
        AttendanceDailies a
    LEFT JOIN 
        Employees e ON e.Id = a.EmployeeId
    LEFT JOIN 
        Branches b ON b.Id = a.BranchId
		    LEFT JOIN Departments d ON d.Id =a.DepartmentId
    LEFT JOIN Designations ds ON ds.Id = e.CurrentDesignationId
    WHERE 
        a.AttendanceDate BETWEEN ''' + CAST(@startDate AS CHAR(10)) + ''' AND ''' + CAST(@endDate AS CHAR(10)) + '''
    GROUP BY
        EmployeeId, e.FirstName, e.LastName, b.[Name],  d.Name , ds.Name 
    ORDER BY 				  
        EmployeeId;';
		print(@SQL)
    EXEC(@SQL);
END;



