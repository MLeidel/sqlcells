sqlcells
d1: testfiles/titanic3.xls
SQL
# People that shared a cabin on the Titanic
# where not all from the same cabin survived.

select cabin, sum(survived) lived, count(name) count from d1
	group by cabin
	having lived < count
OUTPUT
out.xlsx
LAUNCH
