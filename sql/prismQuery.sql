select distinct   
              w.Sku,   
              w.Product AS [Name],
              w.PL AS [PL],
              i.NodeOID AS [OID],
              w.Product_id AS [Item_Id],
              w.[Small Series],
              w.[Big Series]
from     
              Work_FlatCategorization(nolock)as w,
              Items(nolock) as i
where    
              w.Product_id=i.ItemId
              and w.Sku in ('')
