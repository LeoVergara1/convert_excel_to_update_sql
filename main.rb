require 'rubyXL'

["AGSmigrationRev.xlsx","CHImigrationRev.xlsx", "TLAmigrationRev.xlsx", "TOLmigrationRev.xlsx", "VIRmigrationRev.xlsx"].each do |name|
  workbook = RubyXL::Parser.parse(name)
  worksheet = workbook[0]
  file = File.open("update_migration_#{name}.txt", "w+")
  worksheet.drop(1).each do |cells|
    file.puts "UPDATE NOMINA.PRN_TEACHER_PROFILE SET CAMPUS_BASE = '#{cells[6].value}', DIVISION_BASE = '#{cells[7].value}' WHERE PIDM = '#{cells[1].value}';\n"
  end
end
#Columnas Diferentes -_-
["SLPmigrationRev.xlsx","CMXmigrationRev.xlsx", "EMPmigrationRev.xlsx", "GDLmigrationRev.xlsx", "LEOmigrationRev.xlsx", "MERmigrationRev.xlsx","PCHmigrationRev.xlsx", "QROmigrationRev.xlsx", "REFmigrationRev.xlsx"].each do |name|
  workbook = RubyXL::Parser.parse(name)
  worksheet = workbook[0]
  file = File.open("update_migration_#{name}.txt", "w+")
  worksheet.drop(1).each do |cells|
    file.puts "UPDATE NOMINA.PRN_TEACHER_PROFILE SET CAMPUS_BASE = '#{cells[5].value}', DIVISION_BASE = '#{cells[6].value}' WHERE PIDM = '#{cells[1].value}';\n"
  end
end
