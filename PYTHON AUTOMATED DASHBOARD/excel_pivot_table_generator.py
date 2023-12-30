import pandas as pd

df=pd.read_excel('TWO_YEARS_SALES.xlsx')
dp=pd.read_csv('products.csv')
dg=pd.read_csv('customers.csv')
df=df[['order_date','sales_id','product_id','state','customer_id','quantity','sales']]
dp=dp[['product_id','product_name','product_type','size','colour','price']]
dg=dg[['customer_id','gender','age','state']]

#print(inner)



inner=pd.merge(df,dg)    


final=pd.merge(inner,dp)



#print(final)
final.to_excel('merged_updated.xlsx','merged_updated', startrow=4)
pivot_table_1=final.pivot_table(index='product_type',columns='colour',values='sales',aggfunc='sum')

pivot_table_2=final.pivot_table(index='product_type',columns='size',values='quantity',aggfunc='sum')

pivot_table_3=final.pivot_table(index='product_type',columns='product_name',values='quantity',aggfunc='sum')

pivot_table_4=final.pivot_table(index='product_type',columns='state',values='quantity',aggfunc='sum')

pivot_table_5=final.pivot_table(index='state',columns='order_date',values='quantity',aggfunc='sum')

pivot_table_6=final.pivot_table(index='order_date',columns='state',values='sales',aggfunc='sum')

pivot_table_7=final.pivot_table(index='gender',columns='colour',values='quantity',aggfunc='sum')

pivot_table_8=final.pivot_table(index='product_type',columns='gender',values='quantity',aggfunc='sum')

pivot_table_9=final.pivot_table(index='product_type',columns='gender',values='sales',aggfunc='sum')

 
 
#final=pd.ExcelWriter('pivot_tables.xlsx')


#pivot_table_1.to_excel('pivot_tables.xlsx','sales_colours', startrow=4)

pivot_tables=pd.ExcelWriter('pivot_tables.xlsx')

pivot_table_1.to_excel(pivot_tables,sheet_name='sales_colour',index=True)

pivot_table_2.to_excel(pivot_tables,sheet_name='sales_size',index=True)

pivot_table_3.to_excel(pivot_tables,sheet_name='sales_product',index=True)

pivot_table_4.to_excel(pivot_tables,sheet_name='units_state',index=True)

pivot_table_5.to_excel(pivot_tables,sheet_name='units_date',index=True)

pivot_table_6.to_excel(pivot_tables,sheet_name='sales_state',index=True)

pivot_table_7.to_excel(pivot_tables,sheet_name='orders_gender',index=True)

pivot_table_8.to_excel(pivot_tables,sheet_name='product_gender',index=True)

pivot_table_9.to_excel(pivot_tables,sheet_name='spending_gender',index=True)

pivot_tables.close()

#print(pivot_table_1)

#print(pivot_table_2)

#print(pivot_table_3)

#print(pivot_table_4)

#print(pivot_table_5)

#print(pivot_table_6)

#print(dg)


