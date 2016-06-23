# Jupyter - Join unique user IDs from two tables
SELECT DistictUUsersID.user_guid AS userid, d.breed, d.weight, count(*) AS numrows
FROM (SELECT DISTINCT u.user_guid
      FROM users u) AS DistictUUsersID
LEFT JOIN dogs d
ON DistictUUsersID.user_guid = d.user_guid
GROUP BY DistictUUsersID.user_guid
HAVING numrows > 10
ORDER BY numrows DESC;