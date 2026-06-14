# initialize the database, clean all data and only one record root in user table.
echo "use the environment file payroll_local_sqlite.sh"

cd ..
source payroll_local_sqlite.sh 


# pupulate the payroll_details in sqlite
cd /home/richard/shared/jianglei/payroll/payroll_excel_processing
python batch_process.py

# preprocessing for date column ,please all 全角字符
./cleansing_data_dbcs_handling_step0.py   
# delete the outlieres rows
# remove those rows with blank values only (process, ammount, etc)
./cleansing_outliers_step1.py   

# ffile the date value
./cleansing_outliers_step2.py   

# delete the rows contains '合计
./cleansing_outliers_step3.py   

# update the 4 月 5月
./cleansing_outliers_step4.py   

# update the date values
./cleansing_date_handling_step5.py 

# Complex/mixed - The tricky ones 
python3 cleansing_date_handling_step6.py

# remaining ~ handling
python3 cleansing_date_handling_step7.py

# misc remaining issues handled.
python cleansing_misc_step8.py

# remove yy,m / m prefix patterns (e.g. '14,6,3' -> '3', '6,1' -> '1')
./cleansing_date_handling_step9.py

# Step 10: 清理 pandas astype(str) 引入的字面量 'None' 占位符 (代码/客户名称/备注/工序/型号/工序全名, 共 ~863K 行)
python3 cleansing_none_cleanup_step10.py
