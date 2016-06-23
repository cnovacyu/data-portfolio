# Teradata - SKU with greatest sales increase from Nov to Dec
SELECT sku, (Dec_Sales - Nov_Sales) AS Sales_Diff
FROM (SELECT sku, Sum(CASE
      WHEN EXTRACT(MONTH FROM saledate) = 12 THEN (quantity * sprice)                                   
      END) AS Dec_Sales,
        
      Sum(CASE
      WHEN EXTRACT(MONTH FROM saledate) = 11 THEN (quantity * sprice)
      END) AS Nov_Sales
      FROM trnsact
GROUP BY sku) AS SalesMonthAssigned
ORDER BY Sales_Diff DESC;