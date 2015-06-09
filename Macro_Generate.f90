!  Macro_Generate.f90
!
!  FUNCTIONS:
!  Macro_Generate - Entry point of console application.
!

!****************************************************************************
!
!  PROGRAM: Macro_Generate
!
!  PURPOSE:  Entry point for the console application.
!
!****************************************************************************
module variables
	integer ::fileid=11
	integer ::fileid1=9
	character(len=30) ::filename="STATION_ID.txt"
	character(len=30) ::filename2="Macro.txt"
    integer i,j,k
    character(len=60) ::file_path
    integer station_num, station_id
    character(len=30) ::station_id_string
    character(len=30) ::month
    character(len=60) ::new_file_name
    character(len=10) ::row_num
    integer station1,station2
end module

    program Macro_Generate
    use variables
    implicit none
	open (fileid, file=filename)
	open (fileid1, file=filename2)
    write(*,*) "station to station"
    read(*,*) station1,station2
	write (fileid1,*) "Sub Merge_Station_files()"
    write (fileid1, *) "Application.CutCopyMode = False"
    write (fileid1, *) "Application.DisplayAlerts = False"
    !write (fileid1, *) 'ChDir "F:\test\sample_files"'

	!read (fileid, "(A30)") file_path
	read (fileid, *) station_num
	do i = 1,station_num

		read (fileid, *) station_id
        if ((i.ge.station1).and.(i.lt.station2)) then
        write (station_id_string,"(I8)") station_id
		do j = 1, 12
		write (month, "(I2)") j
        write (row_num,"(I8)") 289*(j-1)+1
		new_file_name='"C:\PeMS_DATA\I_15_Flow_South\xls\'//trim(adjustl(station_id_string))//'_'//trim(adjustl(month))//'.xls"'
        write (*,*) new_file_name
        write (fileid1, *) "Workbooks.Open Filename:="//new_file_name
        write (fileid1, *) 'Windows("'//trim(adjustl(station_id_string))//'_'//trim(adjustl(month))//'.xls").Activate'
        write(fileid1,"(A200)",advance='no') 'ActiveWorkbook.SaveAs Filename:='//trim(adjustl(new_file_name))//', FileFormat:=xlExcel8, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False'
        write (fileid1,*)
        write (fileid1,*) "ActiveWindow.Close"
        enddo
        endif
    end do
    write(fileid1,*) "End Sub"
	close (fileid)
	close (fileid1)
    end program Macro_Generate
