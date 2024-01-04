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
              and w.Sku in ('6P7K7EC', '75W47UP', '7D0J1UP', '7N6H3EC', '7N7S8EC',
    '7N7T8EC', '7X6W1UC', '7X6Y7UC', '7Y837UP', '7Z848UP',
    '829U7PC', '83K92UP', '85F65EP', '870N6UC', '870N7UC',
    '872L7UC', '874U9PC', '877Y2UC', '87F17UP', '87F18UP',
    '87F37UP', '87F65UP', '87F66UP', '87G15UP', '885X6UP',
    '886A4UC', '886F3UP', '886F4PC', '886K3PC', '889S6UP',
    '889S7UP', '88F54PC', '88F87PC', '88Q66UP', '88S79EC',
    '88W76UC', '88Y07UP', '893P0EC', '895D2UP', '896D3UP',
    '896F0UP', '897B7UP', '898C3UP', '898Q0UP', '898R1UP',
    '8B5H4UP', '8B7V5EC', '8B815EC', '8B843EC', '8C034EC',
    '8C0N1EC', '8C0Z9EC', '8C106EC', '8C149EC', '8C150EC',
    '8C1J9EC', '8C1R5UC', '8C275EC', '8C278EC', '8C2B3EC',
    '8C7N7EC', '8D5H4UP', '8D6J7UC', '8D8T2UC', '8D8U2EC',
    '8D8Z8EC', '8D903EC', '8D904EC', '8D905EC', '8D907EC',
    '8D918EC', '8D921EC', '8E4L7UC', '8E9L2UC', '8F1D0EC',
    '8F1K2EC', '8F2G4EC', '8F2H9UP', '8F2M0UP', '8F4A7UC',
    '8F6G7UP', '8F6V0UP', '8F704UP', '8G035UP', '8G946UC',
    '8G9A5UC', '8G9F6EP', '8G9G8UP', '8H5C0UP', '8H5R4UP',
    '8H7X1UC', '8J0P0EC', '8J8Q7UP', '8R7Q4UP', '8R7T4EC',
    '8R8D4UP', '8R8D5UP', '8T538EC')
