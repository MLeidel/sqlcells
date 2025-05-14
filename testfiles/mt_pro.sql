sqlcells
d1: testfiles/mt_cost.xlsx
d2: testfiles/mt_sales.xlsx
SQL
select d1.item, sum(amount - cost) Profit from d1, d2
	where d1.item = d2.item
	group by d1.item
OUTPUT
/home/ml/apps/python/projects/sqlcells/out.xlsx
LAUNCH
