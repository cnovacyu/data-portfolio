# Teradata - Vendor with most distinct SKUs only in transaction table
SELECT s.vendor, COUNT(DISTINCT s.sku) AS NumSKUs
FROM skuinfo s
WHERE EXISTS (SELECT t.sku
              FROM trnsact t
              WHERE s.sku = t.sku)
GROUP BY s.vendor
ORDER BY NumSKUs DESC;