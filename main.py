from devpysheet import WorkSheet

service_file = 'pysheets-demo-362905-6512a2f27ed2.json'
worksheet = WorkSheet(service_file=service_file, spread_sheet_title='py_sheets_demo', work_sheet_title='demo1')
addr = worksheet.get_address_with_unique_id(str(7540))
worksheet.set_absent(addr)
