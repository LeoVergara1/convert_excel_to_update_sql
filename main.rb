require 'rubyXL'

["AGSmigrationRev","CHImigrationRev", "TLAmigrationRev", "TOLmigrationRev", "VIRmigrationRev"].each do |name|
  workbook = RubyXL::Parser.parse("#{name}.xlsx")
  worksheet = workbook[0]
  file = File.open("update_migration_#{name}.sql", "w+")
  worksheet.drop(1).each do |cells|
    file.puts "UPDATE NOMINA.PRN_TEACHER_PROFILE SET CAMPUS_BASE = '#{cells[6].value}', DIVISION_BASE = '#{cells[7].value}', USERNAME = 'migrationSql' WHERE PIDM = '#{cells[1].value}';\n"
  end
  p "#{name}.xlsx"
end
#Columnas Diferentes -_-
["SLPmigrationRev","CMXmigrationRev", "EMPmigrationRev", "GDLmigrationRev", "LEOmigrationRev", "MERmigrationRev","PCHmigrationRev", "QROmigrationRev", "REFmigrationRev"].each do |name|
  workbook = RubyXL::Parser.parse("#{name}.xlsx")
  worksheet = workbook[0]
  file = File.open("update_migration_#{name}.sql", "w+")
  worksheet.drop(1).each do |cells|
    file.puts "UPDATE NOMINA.PRN_TEACHER_PROFILE SET CAMPUS_BASE = '#{cells[5].value}', DIVISION_BASE = '#{cells[6].value}', USERNAME = 'migrationSql' WHERE PIDM = '#{cells[1].value}';\n"
  end
  p "#{name}.xlsx"
end
