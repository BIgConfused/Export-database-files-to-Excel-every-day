# Export-database-files-to-Excel-every-day
可每天定时将指定数据库所有表及表中数据备份到Excel中，Excel名为当前天时间戳，表名为sheet名，字段为标题第一行,数据在对应字段列
  数据库配置修改jdbc.properties,然后直接修改ScheRun里的方法参数，详细参数说明请看POIDbToExcel
