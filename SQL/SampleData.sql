--#connection "sqlite:~\..\Data\WorldData.db"

--#table Regions
--#format table Regions for "Contains([Short Name], 'Euro')" with BackColor=Yellow
--#format table Regions column [Table Name] with Condition=Equal, Value1='OECD members', BackColor=Green, ApplyToRow=true
select ID, [Country Code], [Short Name], [Table Name], [Long Name],
	[2-alpha code], [Currency Unit], [Special Notes], [WB-2 code], 
	[National accounts base year],
	[Balance of Payments Manual in use], [External debt Reporting status],
	[Latest agricultural census], [Latest industrial data], [Latest trade data]
from Countries
where ifnull(Region, '') = '';

--#table Countries with GroupBy="Region"
--#format table Countries for "Contains([Region], 'Europe')" with BackColor=Yellow
select ID, [Country Code], [Short Name], [Table Name], [Long Name],
	[2-alpha code], [Currency Unit], [Special Notes], [Region], 
	[Income group], [WB-2 code], [National accounts base year],
	[National accounts reference year], [SNA price valuation],
	[Lending category], [Other groups], [System of National Accounts],
	[Alternative conversion factor], [PPP survey year],
	[Balance of Payments Manual in use], [External debt Reporting status],
	[System of trade], [Government Accounting concept],
	[IMF data dissemination standard], [Latest population census],
	[Latest household survey], [Source of most recent Income and expenditure data],
	[Vital registration complete], [Latest agricultural census],
	[Latest industrial data], [Latest trade data] 
from Countries
where Region > '';

/*#
table Series with GroupBy="Topic Group", OrderBy="Topic Part desc";
format table Series column [Indicator Name] for "Contains([Indicator Name], 'Agricultural')" with BackColor=Green, ForeColor=White;
format table Series column Topic for "Contains([Topic], 'Economic')" with BackColor=LightGreen;
computed column [Topic Part] string in Series = "REGEXMATCH([Topic], '(?<=:).*')";
#*/
select ID, [Series Code], RegexMatch([Series Code], '[^\.]+') as [Series Group],
	RegexMatch([Series Code], '[^\.]+', 1) as [Series SubGroup 1],
	RegexMatch([Series Code], '[^\.]+', 2) as [Series SubGroup 2],
	RegexMatch([Series Code], '[^\.]+', 3) as [Series SubGroup 3],
	RegexMatch([Series Code], '[^\.]+', 4) as [Series SubGroup 4],
	[Topic], RegexMatch([Topic], '[^:]*') as [Topic Group],
	[Indicator Name], [Short definition],
	[Long definition], [Unit of measure], [Periodicity], [Base period],
	[Other notes], [Aggregation method], [Limitations and exceptions], 
	[Notes from original source], [General comments], [Source],
	[Statistical concept and methodology], [Development relevance],
	[Related source links], [Other web links], [Related indicators],
	[License Type]
from Series;

--#table "Energy use - by Regions"
--#format table "Energy use - by Regions" column MaxValue with IconSet=Signs3
--#format table "Energy use - by Regions" column MinValue with IconSet=Arrows3
--#relation Rel_Regions_EnergyUse [Regions] ([Short Name]) - [Energy use - by Regions] ([Region]) 
select c.Region, 
	min(pp.Value) as MinValue,
	max(pp.Value) as MaxValue,
	avg(pp.Value) as AverageValue,
	median(pp.Value) as MedianValue,
	stdev(pp.Value) as StdDev
from [EG.GDP.PUSE.KO.PP] pp
join Countries c on c.[Country Code] = pp.[Country Code]
where Year = 2014 and Value is not null and c.Region > ''
group by c.Region
order by Region;

--#table "Energy use - by Income"
--#format table [Energy use - by Income] column [AverageValue] with Rule=DataBar, BackColor='LightGreen', BackColor2='LightGray', Gradient=Horizontal
--#relation Rel_EnergyUse_Countries [Energy use - by Income] ([Income Group]) - [Countries] ([Income Group])
select c.[Income Group], 
	min(pp.Value) as MinValue,
	max(pp.Value) as MaxValue,
	avg(pp.Value) as AverageValue,
	median(pp.Value) as MedianValue,
	stdev(pp.Value) as StdDev
from [EG.GDP.PUSE.KO.PP] pp
join Countries c on c.[Country Code] = pp.[Country Code]
where Year = 2014 and Value is not null and c.Region > ''
group by c.[Income Group]
order by case [Income Group] 
		when 'Low income' then 1
		when 'Lower middle income' then 2
		when 'Upper middle income' then 3
		when 'High income' then 4 end;

--#table "Energy use"
--#format table "Energy use" column [Income Rank] with IconSet=Rating5
--#format table "Energy use" column Value with IconSet=Arrows5
--#format table "Energy use" column Value with Rule=AboveAverage, BackColor=LightGreen
--#format table "Energy use" column Value with Rule=BelowAverage, BackColor=Red, ForeColor=White
select pp.ID, pp.[Country Code], c.[Table Name] as Country, c.Region, 
	c.[Income Group], case [Income Group] 
		when 'Low income' then 1
		when 'Lower middle income' then 2
		when 'Upper middle income' then 3
		when 'High income' then 4 end as [Income Rank],
	pp.Year, pp.Value
from [EG.GDP.PUSE.KO.PP] pp
join Countries c on c.[Country Code] = pp.[Country Code]
where Year = 2014 and Value is not null and c.Region > ''
order by Country;
