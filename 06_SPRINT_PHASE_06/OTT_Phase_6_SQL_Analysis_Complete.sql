/*
===============================================================================
PHASE 6 — SQL QUERY SCRIPTS & STRUCTURED BUSINESS INSIGHTS
Project: Analyse OTT Platform Viewing Patterns to Support Data-Driven Content Decisions
Tool: MySQL 8.0+ / MySQL Workbench

SOURCE DATA
-----------
1. user_profile_cleaned.xlsx             -> users
2. content_library_cleaned.xlsx          -> content
3. viewing_activity_cleaned.xlsx         -> viewing_activity
4. ratings_feedback_cleaned.xlsx        -> ratings_feedback
5. subscription_retention_cleaned.xlsx  -> subscriptions

IMPORTANT DATA RELATIONSHIP
---------------------------
The content library uses show_id values such as S00001, while viewing and
ratings files use Content_ID values such as C00001. The numeric suffix is the
common content key. Therefore, content joins use:

CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
=
CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)

User_ID is consistent across customer, viewing, ratings and subscription data.

PHASE 6 PURPOSE
---------------
Use SQL to answer the same six business objectives used in Phases 3–5:
O1 Content Performance & Viewer Engagement
O2 Platform & Content Portfolio Performance
O3 Customer Viewing Behaviour
O4 Device & Time-Based Viewing Patterns
O5 Subscription Retention & Churn
O6 Ratings, Feedback & Content Success

This script contains:
- Database creation
- Table creation
- Data validation
- 30 objective-wise business queries
- JOINs, filters and aggregations
- CASE expressions
- CTE/subquery examples
- Derived/summarized tables
- Two recommended complex queries for the Phase 6 video
- Business insight templates

DATA IMPORT NOTE
----------------
MySQL Workbench does not directly execute an Excel .xlsx file through SQL.
Recommended workflow:
1. Convert each cleaned Excel sheet to CSV, OR use MySQL Workbench's
   Table Data Import Wizard.
2. Import into the five tables created below.
3. Run the validation section.
4. Run O1–O6 queries.
Do not change the source data during SQL analysis.

===============================================================================
1. DATABASE
===============================================================================
*/

DROP DATABASE IF EXISTS OTT_Analytics;
CREATE DATABASE OTT_Analytics;
USE OTT_Analytics;

/*
===============================================================================
2. TABLE CREATION
===============================================================================
*/

CREATE TABLE users (
    User_ID VARCHAR(20) PRIMARY KEY,
    Gender VARCHAR(20),
    Age_Group VARCHAR(20),
    Region VARCHAR(30),
    Subscription_Type VARCHAR(30),
    Subscription_Start_Date DATE,
    Subscription_End_Date DATE,
    Device_Type VARCHAR(30),
    Engagement_Level VARCHAR(20)
);

CREATE TABLE content (
    show_id VARCHAR(20) PRIMARY KEY,
    `type` VARCHAR(20),
    title VARCHAR(500),
    director TEXT,
    `cast` TEXT,
    country TEXT,
    release_year INT,
    rating VARCHAR(50),
    duration VARCHAR(50),
    listed_in TEXT,
    description TEXT,
    Platform VARCHAR(50),
    Duration_Minutes DECIMAL(10,2),
    Content_Age_Years INT,
    Numeric_Rating INT,
    Popularity_Score DECIMAL(10,2),
    Engagement_Proxy DECIMAL(10,2),
    Is_Recent TINYINT,
    Genre_Count INT
);

CREATE TABLE viewing_activity (
    Session_ID VARCHAR(30) PRIMARY KEY,
    User_ID VARCHAR(20),
    Content_ID VARCHAR(20),
    View_Date DATE,
    Watch_Duration_Minutes INT,
    Completion_Percentage INT,
    Paused_Times INT,
    Rewatched_Flag VARCHAR(5),
    Time_of_Day VARCHAR(20),
    Device_Type VARCHAR(30)
);

CREATE TABLE ratings_feedback (
    Rating_ID BIGINT AUTO_INCREMENT PRIMARY KEY,
    User_ID VARCHAR(20),
    Content_ID VARCHAR(20),
    Rating INT,
    Liked_Flag VARCHAR(5),
    Feedback_Category VARCHAR(30)
);

CREATE TABLE subscriptions (
    User_ID VARCHAR(20) PRIMARY KEY,
    Subscription_Type VARCHAR(30),
    Monthly_Fee INT,
    Renewal_Status VARCHAR(30),
    Churn_Flag TINYINT
);

/*
===============================================================================
3. DATA IMPORT
===============================================================================

OPTION A — MYSQL WORKBENCH TABLE DATA IMPORT WIZARD
---------------------------------------------------
Use:
Server -> Data Import / Table Data Import Wizard

Import:
user_profile_cleaned.xlsx             -> users
content_library_cleaned.xlsx          -> content
viewing_activity_cleaned.xlsx         -> viewing_activity
ratings_feedback_cleaned.xlsx         -> ratings_feedback
subscription_retention_cleaned.xlsx   -> subscriptions

If converting to CSV, preserve the header row and use UTF-8 CSV.

OPTION B — LOAD DATA LOCAL INFILE
---------------------------------
Export each Excel file as CSV first, then update the paths below.

Example:

LOAD DATA LOCAL INFILE 'C:/YOUR_PATH/user_profile_cleaned.csv'
INTO TABLE users
FIELDS TERMINATED BY ','
ENCLOSED BY '"'
LINES TERMINATED BY '\n'
IGNORE 1 ROWS;

Repeat for the other four tables.

Do not run the example until the actual CSV path is available.
*/

/*
===============================================================================
4. DATA VALIDATION
===============================================================================
*/

-- 4.1 Row counts
SELECT 'users' AS Table_Name, COUNT(*) AS Row_Count FROM users
UNION ALL
SELECT 'content', COUNT(*) FROM content
UNION ALL
SELECT 'viewing_activity', COUNT(*) FROM viewing_activity
UNION ALL
SELECT 'ratings_feedback', COUNT(*) FROM ratings_feedback
UNION ALL
SELECT 'subscriptions', COUNT(*) FROM subscriptions;

-- Expected counts from the cleaned datasets:
-- users              = 2000
-- content            = 17797
-- viewing_activity   = 30000
-- ratings_feedback   = 12000
-- subscriptions      = 2000

-- 4.2 Duplicate user IDs
SELECT User_ID, COUNT(*) AS Duplicate_Count
FROM users
GROUP BY User_ID
HAVING COUNT(*) > 1;

-- 4.3 Duplicate session IDs
SELECT Session_ID, COUNT(*) AS Duplicate_Count
FROM viewing_activity
GROUP BY Session_ID
HAVING COUNT(*) > 1;

-- 4.4 Rating range
SELECT MIN(Rating) AS Min_Rating, MAX(Rating) AS Max_Rating
FROM ratings_feedback;

-- 4.5 Completion range
SELECT MIN(Completion_Percentage) AS Min_Completion,
       MAX(Completion_Percentage) AS Max_Completion
FROM viewing_activity;

-- 4.6 Churn values
SELECT Churn_Flag, COUNT(*) AS Customers
FROM subscriptions
GROUP BY Churn_Flag;

-- 4.7 Renewal status
SELECT Renewal_Status, COUNT(*) AS Customers
FROM subscriptions
GROUP BY Renewal_Status;

/*
===============================================================================
5. RELATIONSHIP / JOIN VALIDATION
===============================================================================
*/

-- 5.1 Viewing records whose User_ID exists in users
SELECT COUNT(*) AS Valid_User_Viewing_Rows
FROM viewing_activity v
JOIN users u
  ON v.User_ID = u.User_ID;

-- 5.2 Viewing records whose content exists in content
SELECT COUNT(*) AS Valid_Content_Viewing_Rows
FROM viewing_activity v
JOIN content c
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED);

-- 5.3 Ratings connected to users
SELECT COUNT(*) AS Valid_User_Rating_Rows
FROM ratings_feedback r
JOIN users u
  ON r.User_ID = u.User_ID;

-- 5.4 Ratings connected to content
SELECT COUNT(*) AS Valid_Content_Rating_Rows
FROM ratings_feedback r
JOIN content c
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(r.Content_ID, 2) AS UNSIGNED);

/*
===============================================================================
O1 — CONTENT PERFORMANCE & VIEWER ENGAGEMENT
30 BUSINESS QUESTIONS START BELOW
===============================================================================
*/

-- Q1.1 Which genres have the largest content catalogue?
SELECT
    TRIM(SUBSTRING_INDEX(listed_in, ',', 1)) AS Primary_Genre,
    COUNT(*) AS Title_Count
FROM content
GROUP BY Primary_Genre
ORDER BY Title_Count DESC
LIMIT 10;

-- Q1.2 Which genres have the highest average completion?
SELECT
    TRIM(SUBSTRING_INDEX(c.listed_in, ',', 1)) AS Primary_Genre,
    COUNT(v.Session_ID) AS Viewing_Sessions,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM viewing_activity v
JOIN content c
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
GROUP BY Primary_Genre
HAVING COUNT(v.Session_ID) >= 20
ORDER BY Avg_Completion DESC
LIMIT 10;

-- Q1.3 Do Movies or TV Shows have higher viewing performance?
SELECT
    c.`type` AS Content_Type,
    COUNT(v.Session_ID) AS Total_Sessions,
    ROUND(AVG(v.Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM viewing_activity v
JOIN content c
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
GROUP BY c.`type`
ORDER BY Avg_Completion DESC;

-- Q1.4 How does content age relate to popularity?
SELECT
    CASE
        WHEN Content_Age_Years <= 2 THEN '0-2 Years'
        WHEN Content_Age_Years <= 5 THEN '3-5 Years'
        WHEN Content_Age_Years <= 10 THEN '6-10 Years'
        ELSE '10+ Years'
    END AS Content_Age_Group,
    COUNT(*) AS Titles,
    ROUND(AVG(Popularity_Score), 2) AS Avg_Popularity,
    ROUND(AVG(Engagement_Proxy), 2) AS Avg_Engagement
FROM content
GROUP BY Content_Age_Group
ORDER BY Avg_Popularity DESC;

-- Q1.5 Which titles have the strongest popularity and engagement?
SELECT
    show_id AS Content_ID,
    title,
    Platform,
    `type` AS Content_Type,
    Popularity_Score,
    Engagement_Proxy
FROM content
ORDER BY Popularity_Score DESC, Engagement_Proxy DESC
LIMIT 10;

/*
===============================================================================
O2 — PLATFORM & CONTENT PORTFOLIO PERFORMANCE
===============================================================================
*/

-- Q2.1 Which platform has the largest content catalogue?
SELECT
    Platform,
    COUNT(*) AS Total_Titles
FROM content
GROUP BY Platform
ORDER BY Total_Titles DESC;

-- Q2.2 Which platform has the highest viewing activity?
SELECT
    c.Platform,
    COUNT(v.Session_ID) AS Total_Viewing_Sessions
FROM content c
LEFT JOIN viewing_activity v
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
GROUP BY c.Platform
ORDER BY Total_Viewing_Sessions DESC;

-- Q2.3 Which platform has the highest average completion?
SELECT
    c.Platform,
    COUNT(v.Session_ID) AS Total_Sessions,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM content c
JOIN viewing_activity v
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
GROUP BY c.Platform
ORDER BY Avg_Completion DESC;

-- Q2.4 Which platform has the highest average watch duration?
SELECT
    c.Platform,
    ROUND(AVG(v.Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM content c
JOIN viewing_activity v
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
GROUP BY c.Platform
ORDER BY Avg_Watch_Duration DESC;

-- Q2.5 Which genres perform best on each platform?
WITH PlatformGenre AS (
    SELECT
        c.Platform,
        TRIM(SUBSTRING_INDEX(c.listed_in, ',', 1)) AS Primary_Genre,
        COUNT(v.Session_ID) AS Sessions,
        AVG(v.Completion_Percentage) AS Avg_Completion
    FROM content c
    JOIN viewing_activity v
      ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
       = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
    GROUP BY c.Platform, Primary_Genre
)
SELECT
    Platform,
    Primary_Genre,
    Sessions,
    ROUND(Avg_Completion, 2) AS Avg_Completion
FROM PlatformGenre
WHERE Sessions >= 20
ORDER BY Platform, Avg_Completion DESC;

/*
===============================================================================
O3 — CUSTOMER VIEWING BEHAVIOUR
===============================================================================
*/

-- Q3.1 Which age group has the highest viewing activity?
SELECT
    u.Age_Group,
    COUNT(v.Session_ID) AS Viewing_Sessions,
    ROUND(AVG(v.Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM users u
JOIN viewing_activity v
  ON u.User_ID = v.User_ID
GROUP BY u.Age_Group
ORDER BY Viewing_Sessions DESC;

-- Q3.2 How does engagement vary by region?
SELECT
    u.Region,
    u.Engagement_Level,
    COUNT(v.Session_ID) AS Viewing_Sessions
FROM users u
JOIN viewing_activity v
  ON u.User_ID = v.User_ID
GROUP BY u.Region, u.Engagement_Level
ORDER BY u.Region, Viewing_Sessions DESC;

-- Q3.3 How does viewing behaviour differ by gender?
SELECT
    u.Gender,
    COUNT(v.Session_ID) AS Viewing_Sessions,
    ROUND(AVG(v.Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM users u
JOIN viewing_activity v
  ON u.User_ID = v.User_ID
GROUP BY u.Gender
ORDER BY Avg_Completion DESC;

-- Q3.4 Which engagement level has the highest viewing intensity?
SELECT
    u.Engagement_Level,
    COUNT(DISTINCT u.User_ID) AS Customers,
    COUNT(v.Session_ID) AS Viewing_Sessions,
    ROUND(AVG(v.Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM users u
LEFT JOIN viewing_activity v
  ON u.User_ID = v.User_ID
GROUP BY u.Engagement_Level
ORDER BY Avg_Watch_Duration DESC;

-- Q3.5 Which region has the highest average completion?
SELECT
    u.Region,
    COUNT(v.Session_ID) AS Viewing_Sessions,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM users u
JOIN viewing_activity v
  ON u.User_ID = v.User_ID
GROUP BY u.Region
ORDER BY Avg_Completion DESC;

/*
===============================================================================
O4 — DEVICE & TIME-BASED VIEWING PATTERNS
===============================================================================
*/

-- Q4.1 Which device is used most?
SELECT
    Device_Type,
    COUNT(*) AS Viewing_Sessions
FROM viewing_activity
GROUP BY Device_Type
ORDER BY Viewing_Sessions DESC;

-- Q4.2 Which device has the highest average watch duration?
SELECT
    Device_Type,
    COUNT(*) AS Sessions,
    ROUND(AVG(Watch_Duration_Minutes), 2) AS Avg_Watch_Duration
FROM viewing_activity
GROUP BY Device_Type
ORDER BY Avg_Watch_Duration DESC;

-- Q4.3 Which device has the highest completion?
SELECT
    Device_Type,
    ROUND(AVG(Completion_Percentage), 2) AS Avg_Completion
FROM viewing_activity
GROUP BY Device_Type
ORDER BY Avg_Completion DESC;

-- Q4.4 What time of day has the most viewing?
SELECT
    Time_of_Day,
    COUNT(*) AS Viewing_Sessions
FROM viewing_activity
GROUP BY Time_of_Day
ORDER BY Viewing_Sessions DESC;

-- Q4.5 Which device-time combination has the highest watch duration?
SELECT
    Device_Type,
    Time_of_Day,
    COUNT(*) AS Sessions,
    ROUND(AVG(Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(Completion_Percentage), 2) AS Avg_Completion
FROM viewing_activity
GROUP BY Device_Type, Time_of_Day
ORDER BY Avg_Watch_Duration DESC;

/*
===============================================================================
O5 — SUBSCRIPTION RETENTION & CHURN
===============================================================================
*/

-- Q5.1 Which subscription plan has the highest renewal rate?
SELECT
    Subscription_Type,
    COUNT(*) AS Customers,
    SUM(CASE WHEN Renewal_Status = 'Renewed' THEN 1 ELSE 0 END) AS Renewed_Customers,
    ROUND(
        SUM(CASE WHEN Renewal_Status = 'Renewed' THEN 1 ELSE 0 END)
        * 100.0 / COUNT(*),
        2
    ) AS Renewal_Rate_Percent
FROM subscriptions
GROUP BY Subscription_Type
ORDER BY Renewal_Rate_Percent DESC;

-- Q5.2 Which subscription plan has the highest churn rate?
SELECT
    Subscription_Type,
    COUNT(*) AS Customers,
    SUM(Churn_Flag) AS Churned_Customers,
    ROUND(AVG(Churn_Flag) * 100, 2) AS Churn_Rate_Percent
FROM subscriptions
GROUP BY Subscription_Type
ORDER BY Churn_Rate_Percent DESC;

-- Q5.3 Does engagement level relate to churn?
SELECT
    u.Engagement_Level,
    COUNT(DISTINCT u.User_ID) AS Customers,
    SUM(s.Churn_Flag) AS Churned_Customers,
    ROUND(AVG(s.Churn_Flag) * 100, 2) AS Churn_Rate_Percent
FROM users u
JOIN subscriptions s
  ON u.User_ID = s.User_ID
GROUP BY u.Engagement_Level
ORDER BY Churn_Rate_Percent DESC;

-- Q5.4 How do Basic and Premium customers compare?
SELECT
    s.Subscription_Type,
    COUNT(*) AS Customers,
    ROUND(AVG(s.Monthly_Fee), 2) AS Avg_Monthly_Fee,
    ROUND(
        AVG(CASE WHEN s.Renewal_Status = 'Renewed' THEN 1 ELSE 0 END) * 100,
        2
    ) AS Renewal_Rate_Percent,
    ROUND(AVG(s.Churn_Flag) * 100, 2) AS Churn_Rate_Percent
FROM subscriptions s
WHERE s.Subscription_Type IN ('Basic', 'Premium')
GROUP BY s.Subscription_Type
ORDER BY Renewal_Rate_Percent DESC;

-- Q5.5 How does monthly fee relate to churn?
SELECT
    Monthly_Fee,
    COUNT(*) AS Customers,
    ROUND(AVG(Churn_Flag) * 100, 2) AS Churn_Rate_Percent,
    ROUND(
        AVG(CASE WHEN Renewal_Status = 'Renewed' THEN 1 ELSE 0 END) * 100,
        2
    ) AS Renewal_Rate_Percent
FROM subscriptions
GROUP BY Monthly_Fee
ORDER BY Monthly_Fee;

/*
===============================================================================
O6 — RATINGS, FEEDBACK & CONTENT SUCCESS
===============================================================================
*/

-- Q6.1 Which titles have the highest average ratings?
SELECT
    c.title,
    c.Platform,
    COUNT(r.Rating) AS Rating_Count,
    ROUND(AVG(r.Rating), 2) AS Avg_Rating
FROM ratings_feedback r
JOIN content c
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(r.Content_ID, 2) AS UNSIGNED)
GROUP BY c.show_id, c.title, c.Platform
HAVING COUNT(r.Rating) >= 5
ORDER BY Avg_Rating DESC, Rating_Count DESC
LIMIT 10;

-- Q6.2 Which titles receive the most likes?
SELECT
    c.title,
    COUNT(*) AS Total_Feedback,
    SUM(CASE WHEN r.Liked_Flag = 'Y' THEN 1 ELSE 0 END) AS Total_Likes
FROM ratings_feedback r
JOIN content c
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(r.Content_ID, 2) AS UNSIGNED)
GROUP BY c.show_id, c.title
ORDER BY Total_Likes DESC
LIMIT 10;

-- Q6.3 What is the overall feedback distribution?
SELECT
    Feedback_Category,
    COUNT(*) AS Feedback_Count,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM ratings_feedback), 2)
        AS Feedback_Percent
FROM ratings_feedback
GROUP BY Feedback_Category
ORDER BY Feedback_Count DESC;

-- Q6.4 Does higher rating correspond to higher completion?
SELECT
    r.Rating,
    COUNT(*) AS Rated_Viewers,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM ratings_feedback r
JOIN viewing_activity v
  ON r.User_ID = v.User_ID
 AND r.Content_ID = v.Content_ID
GROUP BY r.Rating
ORDER BY r.Rating;

-- Q6.5 Does higher rating correspond to higher rewatch behaviour?
SELECT
    r.Rating,
    COUNT(*) AS Rated_Viewers,
    ROUND(
        AVG(CASE WHEN v.Rewatched_Flag = 'Y' THEN 1 ELSE 0 END) * 100,
        2
    ) AS Rewatch_Rate_Percent
FROM ratings_feedback r
JOIN viewing_activity v
  ON r.User_ID = v.User_ID
 AND r.Content_ID = v.Content_ID
GROUP BY r.Rating
ORDER BY r.Rating;

/*
===============================================================================
6. ADVANCED / COMPLEX SQL — RECOMMENDED FOR PHASE 6 VIDEO
===============================================================================
*/

-- COMPLEX QUERY 1
-- Platform Performance Scorecard
-- Demonstrates LEFT JOIN + aggregation + derived metrics + ORDER BY.

SELECT
    c.Platform,
    COUNT(DISTINCT c.show_id) AS Total_Titles,
    COUNT(v.Session_ID) AS Total_Viewing_Sessions,
    ROUND(AVG(v.Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion,
    ROUND(
        COUNT(v.Session_ID) * 1.0 /
        NULLIF(COUNT(DISTINCT c.show_id), 0),
        2
    ) AS Sessions_Per_Title
FROM content c
LEFT JOIN viewing_activity v
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
GROUP BY c.Platform
ORDER BY Avg_Completion DESC, Avg_Watch_Duration DESC;


-- COMPLEX QUERY 2
-- Engagement vs Churn Scorecard
-- Demonstrates JOIN + CASE + aggregation + derived percentage.

SELECT
    u.Engagement_Level,
    COUNT(DISTINCT u.User_ID) AS Total_Customers,
    SUM(CASE WHEN s.Churn_Flag = 1 THEN 1 ELSE 0 END) AS Churned_Customers,
    SUM(CASE WHEN s.Renewal_Status = 'Renewed' THEN 1 ELSE 0 END)
        AS Renewed_Customers,
    ROUND(
        SUM(CASE WHEN s.Churn_Flag = 1 THEN 1 ELSE 0 END)
        * 100.0 / COUNT(DISTINCT u.User_ID),
        2
    ) AS Churn_Rate_Percent,
    ROUND(
        SUM(CASE WHEN s.Renewal_Status = 'Renewed' THEN 1 ELSE 0 END)
        * 100.0 / COUNT(DISTINCT u.User_ID),
        2
    ) AS Renewal_Rate_Percent
FROM users u
JOIN subscriptions s
  ON u.User_ID = s.User_ID
GROUP BY u.Engagement_Level
ORDER BY Churn_Rate_Percent DESC;


-- COMPLEX QUERY 3
-- Top-performing titles within each platform using CTE + window function.

WITH TitlePerformance AS (
    SELECT
        c.Platform,
        c.show_id,
        c.title,
        COUNT(v.Session_ID) AS Sessions,
        AVG(v.Completion_Percentage) AS Avg_Completion,
        AVG(v.Watch_Duration_Minutes) AS Avg_Watch_Duration
    FROM content c
    JOIN viewing_activity v
      ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
       = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
    GROUP BY c.Platform, c.show_id, c.title
),
RankedTitles AS (
    SELECT
        *,
        DENSE_RANK() OVER (
            PARTITION BY Platform
            ORDER BY Avg_Completion DESC, Sessions DESC
        ) AS Platform_Rank
    FROM TitlePerformance
)
SELECT
    Platform,
    show_id AS Content_ID,
    title,
    Sessions,
    ROUND(Avg_Completion, 2) AS Avg_Completion,
    ROUND(Avg_Watch_Duration, 2) AS Avg_Watch_Duration,
    Platform_Rank
FROM RankedTitles
WHERE Platform_Rank <= 5
ORDER BY Platform, Platform_Rank;


-- COMPLEX QUERY 4
-- Customers with high engagement but churn risk.
SELECT
    u.User_ID,
    u.Engagement_Level,
    s.Subscription_Type,
    s.Monthly_Fee,
    s.Churn_Flag,
    s.Renewal_Status
FROM users u
JOIN subscriptions s
  ON u.User_ID = s.User_ID
WHERE u.Engagement_Level = 'High'
  AND s.Churn_Flag = 1
ORDER BY s.Monthly_Fee DESC;


/*
===============================================================================
7. DERIVED / SUMMARY TABLES
===============================================================================
*/

-- 7.1 Platform summary table
DROP TABLE IF EXISTS platform_summary;

CREATE TABLE platform_summary AS
SELECT
    c.Platform,
    COUNT(DISTINCT c.show_id) AS Total_Titles,
    COUNT(v.Session_ID) AS Total_Viewing_Sessions,
    ROUND(AVG(v.Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM content c
LEFT JOIN viewing_activity v
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
GROUP BY c.Platform;

SELECT * FROM platform_summary
ORDER BY Avg_Completion DESC;


-- 7.2 Subscription summary table
DROP TABLE IF EXISTS subscription_summary;

CREATE TABLE subscription_summary AS
SELECT
    Subscription_Type,
    COUNT(*) AS Total_Customers,
    SUM(Churn_Flag) AS Churned_Customers,
    ROUND(AVG(Monthly_Fee), 2) AS Avg_Monthly_Fee,
    ROUND(
        AVG(CASE WHEN Renewal_Status = 'Renewed' THEN 1 ELSE 0 END) * 100,
        2
    ) AS Renewal_Rate_Percent,
    ROUND(AVG(Churn_Flag) * 100, 2) AS Churn_Rate_Percent
FROM subscriptions
GROUP BY Subscription_Type;

SELECT * FROM subscription_summary
ORDER BY Renewal_Rate_Percent DESC;


-- 7.3 Device-time summary table
DROP TABLE IF EXISTS device_time_summary;

CREATE TABLE device_time_summary AS
SELECT
    Device_Type,
    Time_of_Day,
    COUNT(*) AS Viewing_Sessions,
    ROUND(AVG(Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(Completion_Percentage), 2) AS Avg_Completion
FROM viewing_activity
GROUP BY Device_Type, Time_of_Day;

SELECT *
FROM device_time_summary
ORDER BY Avg_Watch_Duration DESC;


/*
===============================================================================
8. BUSINESS INSIGHT QUERIES
===============================================================================
*/

-- Insight 1: Best platform by completion
SELECT *
FROM platform_summary
ORDER BY Avg_Completion DESC
LIMIT 1;

-- Insight 2: Highest churn plan
SELECT *
FROM subscription_summary
ORDER BY Churn_Rate_Percent DESC
LIMIT 1;

-- Insight 3: Highest renewal plan
SELECT *
FROM subscription_summary
ORDER BY Renewal_Rate_Percent DESC
LIMIT 1;

-- Insight 4: Most-used device
SELECT
    Device_Type,
    COUNT(*) AS Sessions
FROM viewing_activity
GROUP BY Device_Type
ORDER BY Sessions DESC
LIMIT 1;

-- Insight 5: Most active time of day
SELECT
    Time_of_Day,
    COUNT(*) AS Sessions
FROM viewing_activity
GROUP BY Time_of_Day
ORDER BY Sessions DESC
LIMIT 1;


/*
===============================================================================
9. OPTIONAL VIEWS FOR REPORTING
===============================================================================
*/

CREATE OR REPLACE VIEW vw_platform_performance AS
SELECT
    c.Platform,
    COUNT(DISTINCT c.show_id) AS Total_Titles,
    COUNT(v.Session_ID) AS Total_Viewing_Sessions,
    ROUND(AVG(v.Watch_Duration_Minutes), 2) AS Avg_Watch_Duration,
    ROUND(AVG(v.Completion_Percentage), 2) AS Avg_Completion
FROM content c
LEFT JOIN viewing_activity v
  ON CAST(SUBSTRING(c.show_id, 2) AS UNSIGNED)
   = CAST(SUBSTRING(v.Content_ID, 2) AS UNSIGNED)
GROUP BY c.Platform;

CREATE OR REPLACE VIEW vw_customer_retention AS
SELECT
    u.Engagement_Level,
    s.Subscription_Type,
    COUNT(*) AS Customers,
    ROUND(AVG(s.Churn_Flag) * 100, 2) AS Churn_Rate_Percent,
    ROUND(
        AVG(CASE WHEN s.Renewal_Status = 'Renewed' THEN 1 ELSE 0 END) * 100,
        2
    ) AS Renewal_Rate_Percent
FROM users u
JOIN subscriptions s
  ON u.User_ID = s.User_ID
GROUP BY u.Engagement_Level, s.Subscription_Type;

-- Test views
SELECT * FROM vw_platform_performance;
SELECT * FROM vw_customer_retention;


/*
===============================================================================
10. PHASE 6 BUSINESS INSIGHT DOCUMENTATION TEMPLATE
===============================================================================

For every important SQL result, document:

OBJECTIVE:
O1 / O2 / O3 / O4 / O5 / O6

BUSINESS QUESTION:
Write the exact question answered by the query.

SQL TECHNIQUE:
SELECT / WHERE / JOIN / GROUP BY / CASE / CTE / Window Function / Subquery

KEY RESULT:
Write the important value or comparison from the output.

BUSINESS INSIGHT:
Explain what the result means for the OTT business.

RECOMMENDATION:
State a practical action based on the result.

IMPORTANT:
SQL results are descriptive unless supported by formal statistical testing.
Do not claim causation or statistical significance from SQL aggregation alone.

===============================================================================
11. PHASE 6 VIDEO — TWO QUERIES TO EXPLAIN
===============================================================================

Recommended Query 1:
"Platform Performance Scorecard"

Explain:
1. Why content and viewing tables are joined.
2. Why LEFT JOIN is used.
3. Why COUNT(DISTINCT) is used for titles.
4. Why AVG is used for watch duration and completion.
5. What the final ranking means for platform decisions.

Recommended Query 2:
"Engagement vs Churn Scorecard"

Explain:
1. Why users and subscriptions are joined.
2. How CASE identifies churned and renewed customers.
3. How churn rate is calculated.
4. How engagement groups are compared.
5. What the output means for retention strategy.

===============================================================================
12. FINAL PHASE 6 CHECKLIST
===============================================================================

[ ] Database OTT_Analytics created
[ ] Five SQL tables created
[ ] Five cleaned datasets imported
[ ] Row counts verified
[ ] Duplicate / range checks completed
[ ] User joins verified
[ ] Content joins verified
[ ] O1 — 5 questions executed
[ ] O2 — 5 questions executed
[ ] O3 — 5 questions executed
[ ] O4 — 5 questions executed
[ ] O5 — 5 questions executed
[ ] O6 — 5 questions executed
[ ] At least 2 complex queries executed
[ ] Derived summary tables created
[ ] Views created
[ ] Business insights documented
[ ] Screenshots saved for important outputs
[ ] Phase 6 SQL file saved in project folder
[ ] 3–4 minute SQL explanation video prepared

END OF PHASE 6 SQL SCRIPT
===============================================================================
