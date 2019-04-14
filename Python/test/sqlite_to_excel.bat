set db_file=test.db
del "out\%db_file%_test_1.xlsx"
del "out\%db_file%_test_2.xlsx"
cd ..
python export_to_excel.py -v -m 10 -d "test/in/test_db/%db_file%" -o "test/out/%db_file%_test_1.xlsx" -c "test/in/columns.txt" -p "{\"table\":\"test_table\",\"where\":\"id ^< 10 OR id ^> 15\",\"column_list\":[\"id\",\"descr\",\"create_date\"]}"
python export_to_excel.py -d "test/in/test_db/%db_file%" -o "test/out/%db_file%_test_2.xlsx" -c "test/in/columns.txt" -p "{\"table\":\"test_table\",\"where\":\"id ^< 10 OR id ^> 15\",\"column_list\":[\"id\",\"descr\",\"create_date\"]}"
cd test