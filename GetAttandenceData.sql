USE [AVERYLA]
GO
/****** Object:  StoredProcedure [dbo].[GetAttendanceData]    Script Date: 2024-03-03 12:02:26 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author		:	Easin Alif
-- Create date:  2024-02-28
-- Description	:	Datewise Attendance Data
-- Notes			: Columns sequence shound not be change, otherwise excel output data will be mismatch
-- =============================================

-- EXEC GetAttendanceData '2023-02-15',  '2023-02-16'

ALTER PROCEDURE [dbo].[GetAttendanceData]
    @startDate DATE,
    @endDate DATE
AS
BEGIN
    DECLARE @SQL VARCHAR(MAX) = '', @dateRange VARCHAR(MAX) = '';

    SELECT @dateRange += ',' + QUOTENAME(DATEADD(DAY, number, @startDate))
    FROM master..spt_values
    WHERE type = 'P' AND DATEADD(DAY, number, @startDate) <= @endDate;

    SELECT @dateRange = STUFF(@dateRange, 1, 1, '');

    SELECT @SQL = '
SELECT P.* FROM (
    SELECT 
        EmployeeId, 
        e.FirstName +'' ''+ e.LastName AS [EmployeeName], 
        b.[Name] AS [Location], 
        d.Name AS [DepartmentName], 
        ds.Name AS [DesignationName], 
        AttendanceDate, 
        (CASE
            WHEN LeaveName IS NULL THEN Status    
            ELSE LeaveName
        END) AS [Status]
    FROM AttendanceDailies a
    LEFT JOIN Employees e ON e.Id=a.EmployeeId
    LEFT JOIN Branches b ON b.Id=a.BranchId
    LEFT JOIN Departments d ON d.Id =a.DepartmentId
    LEFT JOIN Designations ds ON ds.Id = e.CurrentDesignationId
    WHERE a.AttendanceDate BETWEEN ''' + CAST(@startDate AS CHAR(10)) + ''' AND ''' + CAST(@endDate AS CHAR(10)) + ''' 
) T
PIVOT (MAX([Status]) FOR AttendanceDate IN (' + @dateRange + ')) P
ORDER BY EmployeeId;';

EXEC(@SQL);
END;