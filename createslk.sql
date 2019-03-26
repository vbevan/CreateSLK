--Input = Firstname, Surname, DOB, Gender

--1. Remove Non Alpha Characters from Frstname and Surname
CREATE FUNCTION dbo.RemoveNonAlphaCharacters
(
    @testString VARCHAR(1000)
)
RETURNS VARCHAR(1000)
AS
    BEGIN
        DECLARE @outputString VARCHAR(1000) = '';
        DECLARE @counter INT = 1;
        DECLARE @testChar VARCHAR(1) = '';

        WHILE @counter <= LEN(@testString)
            BEGIN
                SET @testChar = SUBSTRING(@testString, @counter, 1)

                IF ((@testChar
                   BETWEEN CHAR(48) AND CHAR(57))
                    OR (@testChar
                   BETWEEN CHAR(65) AND CHAR(90))
                    OR (@testChar
                   BETWEEN CHAR(97) AND CHAR(122)))
                    BEGIN
                        SET @outputString = @outputString + @testChar
                        SET @counter = @counter + 1
                    END
                ELSE
                    BEGIN
                        SET @counter = @counter + 1
                    END
            END

        RETURN @outputString
    END;
GO
--2. Create Surname SLK Component
CREATE FUNCTION dbo.SLKSurname
(
    @SurName VARCHAR(1000)
)
RETURNS VARCHAR(1000)
AS
    BEGIN
        DECLARE @SurNameAlphaOnly VARCHAR(1000) = dbo.RemoveNonAlphaCharacters(@SurName);
        DECLARE @slkSurname VARCHAR(1000);

        SELECT @slkSurname
            = CASE
                   WHEN LEN(@SurNameAlphaOnly) > 4 THEN
                       SUBSTRING(@SurNameAlphaOnly, 2, 2) + SUBSTRING(@SurNameAlphaOnly, 5, 1)
                   WHEN LEN(@SurNameAlphaOnly) > 2 THEN SUBSTRING(@SurNameAlphaOnly, 2, 2) + '2'
                   WHEN LEN(@SurNameAlphaOnly) > 1 THEN SUBSTRING(@SurNameAlphaOnly, 2, 2) + '22'
                   WHEN LEN(@SurNameAlphaOnly) >= 1 THEN '222'
                   ELSE '999'
              END;

        RETURN UPPER(@slkSurname);
    END;
GO
--3. Create Firstname SLK Component
CREATE FUNCTION [dbo].[SLKFirstName]
(
    @firstName VARCHAR(1000)
)
RETURNS VARCHAR(1000)
AS
    BEGIN
        DECLARE @firstNameAlphaOnly VARCHAR(1000) = dbo.RemoveNonAlphaCharacters(@firstName);
        DECLARE @slkFirstName VARCHAR(1000);

        SELECT @slkFirstName = CASE
                                    WHEN LEN(@firstNameAlphaOnly) > 2 THEN SUBSTRING(@firstNameAlphaOnly, 2, 2)
                                    WHEN LEN(@firstNameAlphaOnly) > 1 THEN SUBSTRING(@firstNameAlphaOnly, 2, 1) + '2'
                                    WHEN LEN(@firstNameAlphaOnly) = 1 THEN '22'
                                    ELSE '99'
                               END;

        RETURN UPPER(@slkFirstName);
    END;
GO	 
--4. Join it all together to create SLK
CREATE FUNCTION dbo.CreateSLK
(
    @SurName VARCHAR(1000)
   ,@firstName VARCHAR(1000)
   ,@DOB DATETIME2
   ,@gender VARCHAR(10)
)
RETURNS VARCHAR(14)
AS
    BEGIN
        DECLARE @slkSurname VARCHAR(3) = dbo.SLKSurname(COALESCE(@SurName, ''));
        DECLARE @slkFirstName VARCHAR(2) = dbo.SLKFirstName(COALESCE(@firstName, ''));
        DECLARE @myDOB VARCHAR(8)
            = REPLACE(CONVERT(VARCHAR(10), COALESCE(@DOB, CAST('01/01/1900' AS DATETIME2)), 103), '/', '');
        DECLARE @myGender VARCHAR(1) = CASE
                                            WHEN UPPER(@gender) IN ( 'M', 'MALE', '1' ) THEN '1'
                                            WHEN UPPER(@gender) IN ( 'F', 'FEMALE', '2' ) THEN '2'
                                            WHEN COALESCE(@gender, '9') = '9' THEN '9'
                                            ELSE '9'
                                       END;

        RETURN @slkSurname + @slkFirstName + @myDOB + @myGender;
    END;
GO