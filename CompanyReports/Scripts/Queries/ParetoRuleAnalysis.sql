with product_wise_sales as (
	select 
		Product.ProductNumber, 
		sum(SalesOrderDetail.LineTotal) as product_sales 
	from Sales.SalesOrderDetail 
	inner join Production.Product on Product.ProductID = SalesOrderDetail.ProductID 
	inner join Sales.SalesOrderHeader on SalesOrderDetail.SalesOrderID = SalesOrderHeader.SalesOrderID 
	where SalesOrderHeader.Status = 5 
	group by Product.ProductNumber 
	),
calc_sales as (
	select 
		product_wise_sales.ProductNumber, 
		product_sales, 
		sum(product_sales) over(order by product_sales desc rows between unbounded preceding and 0 preceding) as running_sales, 
		0.8 * sum(product_sales) over() as total_sales 
	from product_wise_sales
	) 
select * from calc_sales where running_sales <= total_sales