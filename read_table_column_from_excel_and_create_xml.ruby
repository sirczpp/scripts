require "roo"
require "rexml/document"

xls = Roo::Spreadsheet.open("./1.xls")

s1=xls.sheet("11")

doc = REXML::Document.new 
root = doc.add_element("SyncJob")
doc.context[:attribute_quote] = :quote

j=0
table_array = Array.new

for i in(s1.first_row..s1.last_row) do
	if(s1.row(i)[0]!=nil)
		if s1.row(i)[0]!="功能描述"&&s1.row(i)[0]!="列名"
			if s1.row(i)[0]=="表名"
				table_array[j] = i
				j=j+1
			end
		end
	end 
end

puts "\n..............."

table_count = table_array.length

for i in(0..table_count-1) do

	table_row_index = table_array[i]
	pk_row_index = table_row_index + 4
	column_row_index = pk_row_index + 1
	if i!=table_count-1
		last_column_row_index = table_array[i+1]-1
	else
		last_column_row_index = s1.last_row
	end

	table_id = s1.row(table_row_index)[1]
	pk_id=s1.row(pk_row_index)[0]

	table = root.add_element("table")
	table.attributes["table_id"] = table_id
	table.attributes["type"] = "I"

	source = table.add_element("source")
	source.attributes["schema"] = "SIRCZPP"
	source.attributes["name"] = table_id

	destination = table.add_element("destination")
	destination.attributes["schema"] = "SIRCZPP"
	destination.attributes["name"] = table_id
	
	values = table.add_element("values")
	key = values.add_element("key")
	key.attributes["name"] = pk_id
	key.text = pk_id

	sql ='select '

	for j in(column_row_index..last_column_row_index) do
		column_name = s1.row(j)[0]
		column = values.add_element("column")
		column.attributes["name"] = column_name
		column.text = column_name

		if j!=last_column_row_index
			sql=sql+"#{column_name}"+','
		else
			sql=sql+"#{column_name}"+' '
		end	
	end
	t ="where SJC > @ts_lastSucc"

	sql="#{sql}#{t.to_s}"

	source.attributes["sql"] = sql
end

output_string =""
doc.write(:output => output_string, :indent => 2)

#damn it
output_string.gsub!(/&gt;/,'>')

file = File.new("./1.xml","w+")
file.write output_string
