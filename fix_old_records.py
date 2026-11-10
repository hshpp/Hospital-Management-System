import sqlite3

conn = sqlite3.connect('hospital.db')
cursor = conn.cursor()

# نمایش تعداد رکوردها بر اساس invoice_id
cursor.execute('''
    SELECT invoice_id, COUNT(*) as count
    FROM medicine_sales
    GROUP BY invoice_id
    ORDER BY invoice_id DESC
''')
rows = cursor.fetchall()

print('📋 Records grouped by invoice_id:')
for row in rows[:20]:
    print(f'   {row[0]}: {row[1]} items')
conn.close()