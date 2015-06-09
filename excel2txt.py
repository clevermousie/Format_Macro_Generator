import xlrd
input=file('STA_ID.txt')
for line in input.readlines():
	print line
#	line_data=line.split(' ')
#	line_data_0=float(line_data[1])
	line_data_1=int(line)	
	print line_data_1
	month=0
	while month <5:
		month+=1
		month_str=str(month)
		print month_str
		new_name='F:/Dropbox/RESEARCH_PROJECTS/Adverse_Weather_Induced_Delay/Historical_Data/2014_Jan_Apr_DEL60/EXCEL/'+str(line_data_1)+'_'+str(month)+'.xls'
		print new_name
		workbook = xlrd.open_workbook(new_name)
#		workbook = xlrd.open_workbook(new_name)
		print 'new file opened'
		worksheet = workbook.sheet_by_name('Report Data')
		num_rows=worksheet.nrows-1
		num_cells=worksheet.ncols-1
		curr_row=0
		new_name_out='F:/Dropbox/RESEARCH_PROJECTS/Adverse_Weather_Induced_Delay/Historical_Data/2014_Jan_Apr_DEL60/TXT/'+str(line_data_1)+'_'+str(month)+'.txt'
		output=file(new_name_out,'w')
		while curr_row < num_rows:
			curr_row+=1
			row=worksheet.row(curr_row)
			cell_value=worksheet.cell_value(curr_row,0)
			time=(cell_value-int(cell_value))*10000000/34722
			time=int(time)
			hour=int(time/12)
			minute=5*(time-12*int(time/12))
			cell_value2=worksheet.cell_value(curr_row,1)
			print >> output, ' ',int(cell_value)-41274,' ',hour,' ',minute,' ',cell_value2
		output.close()
input.close()
