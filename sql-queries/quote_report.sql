# Pulls all open quotes that have not expired, includes sales rep and region
SELECT "customer"."name" AS "Customer", "quotehed"."quotenum" AS "Quote #",
"quotedtl"."quoteline" AS "Line", "quotedtl"."partnum" AS "Part Num",
"quotedtl"."linedesc" AS "Part Description", "quotehed"."shortchar10" AS "TSR",
CONVERT(VARCHAR(10),"quotehed"."entrydate",101) AS "Date Quoted", CONVERT(VARCHAR(10),"quotehed"."expirationdate",101) AS "Date Expires", 
"customer"."territoryid" AS "Territory",
CASE 
  WHEN "quotehed"."territoryid" = 'US 01' THEN 'Region 1'
  WHEN "quotehed"."territoryid" = 'US 02' THEN 'Region 1'
  WHEN "quotehed"."territoryid" = 'US 03' THEN 'Region 1'
  WHEN "quotehed"."territoryid" = 'US 04' THEN 'Region 1'
  WHEN "quotehed"."territoryid" = 'US 05' THEN 'Region 1'
  WHEN "quotehed"."territoryid" = 'US 06' THEN 'Region 1'
  WHEN "quotehed"."territoryid" = 'CAN 01' THEN 'Region 1'
  WHEN "quotehed"."territoryid" = 'MEX 01' THEN 'Region 1'
  WHEN "quotehed"."territoryid" = 'US 07' THEN 'Region 2'
  WHEN "quotehed"."territoryid" = 'US 08' THEN 'Region 2'
  WHEN "quotehed"."territoryid" = 'US 09' THEN 'Region 2'
  WHEN "quotehed"."territoryid" = 'US 10' THEN 'Region 2'
  WHEN "quotehed"."territoryid" = 'US 11' THEN 'Region 2'
  WHEN "quotehed"."territoryid" = 'US 14' THEN 'Region 2'
  WHEN "quotehed"."territoryid" = 'EUR 01' THEN 'Region 3'
  WHEN "quotehed"."territoryid" = 'EUR 02' THEN 'Region 3'
  ELSE 'Region 4'
  END AS "Region"
FROM "dbo"."quotehed" "quotehed"
  INNER JOIN "dbo"."customer" "customer" ON ("quotehed"."custnum" = "customer"."custnum")
  INNER JOIN "dbo"."quotedtl" "quotedtl" ON ("quotehed"."quotenum" = "quotedtl"."quotenum")
WHERE CONVERT(VARCHAR(10),"quotehed"."expirationdate",101) >= DATEADD(dd, -1, GETDATE()) AND "quotehed"."quoted" = 1
ORDER BY "quotehed"."expirationdate" ASC;
