with Dataset as (
    select 
        CustomerID, 
        SalesOrderID, 
        OrderDate, 
        TotalDue 
    from Sales.SalesOrderHeader 
    inner join Sales.SalesTerritory on SalesTerritory.TerritoryID = SalesOrderHeader.TerritoryID 
    where SalesTerritory.Name = @Country and SalesOrderHeader.Status = 5
    ),
Order_Summary as (
    select
        CustomerID, 
        SalesOrderID, 
        OrderDate, 
        sum(TotalDue) as Total_Sales 
    from Dataset 
    group by CustomerID, SalesOrderID, OrderDate
) 
select
    t1.CustomerID, 
    datediff(day, (select max(OrderDate) from Order_Summary where CustomerID = t1.CustomerID), (select max(OrderDate) from Order_Summary)) as Recency, 
    count(t1.SalesOrderID) as Frequency, 
    sum(t1.Total_Sales) as Monetary, 
    ntile(10) over(order by datediff(day, (select max(OrderDate) from Order_Summary where CustomerID = t1.CustomerID), (select max(OrderDate) from Order_Summary)) desc) as R, 
    ntile(10) over(order by count(t1.SalesOrderID) asc) as F, 
    ntile(10) over(order by sum(t1.Total_Sales) asc) as M 
from Order_Summary t1 
group by t1.CustomerID 
order by 1, 3 desc;
