require 'spreadsheet'
require 'csv'

root_dir = ARGV[0] || 'archivos'
unless Dir.exists? root_dir
  puts "El directorio #{root_dir} no existe"
  exit
end

splitter = ARGV[1] || ','
column = ARGV[2] ? ARGV[2].to_i : 10

files_dir = File.join(root_dir, '**', '*')
puts "Directorio de archivos: #{files_dir}"

total_count = 0

def clearData(data)
  return data.to_s.tr(',', '').strip if data
end

def getValue(data)
  if data
    if data.is_a?(Spreadsheet::Formula) and data.value.is_a?(Float)
      return data.value
    else
      return 0
    end
  else
    return 'ERROR'
  end
end

CSV.open(File.join(root_dir, 'resultado.csv'), "wb") do |csv|

  csv << ['id_rodal', 'id_tipo', 'id_pos', 'id_densidad', 'id_orientacion',
    'poda_rodal', 'raleo_rodal', 'id_uso', 'edad', 'id_estado_muestra',
    'id_cond_muestra', 'id_rechazo', 'fecha', 'id_origen',
    'id_parcela', 'id_distritos', 'id_cond_parcela',
    'latitud', 'longitud',
    'nro_plantas', 'promedio', 'maximo', 'minimo',
    'volumen', 'altura promedio', 'archivo']

  Dir.glob(files_dir) do |xls|
    next if File.extname(xls) != '.xls'
    total_count += 1
    book = Spreadsheet.open xls
    sheet = book.worksheet 0
    # data = File.basename(xls, '.xls').split(splitter).map!(&:strip)

    data = []

    # Leer datos de la cabecera
    data << clearData(sheet.row(1)[1])
    data << clearData(sheet.row(1)[2])
    data << clearData(sheet.row(1)[3])
    data << clearData(sheet.row(1)[4])
    data << clearData(sheet.row(1)[5])
    data << clearData(sheet.row(1)[6])
    data << clearData(sheet.row(1)[7])
    data << clearData(sheet.row(1)[8])
    data << clearData(sheet.row(1)[9])
    data << clearData(sheet.row(3)[1])
    data << clearData(sheet.row(3)[2])
    data << clearData(sheet.row(3)[3])
    data << clearData(sheet.row(3)[4])
    data << clearData(sheet.row(3)[5])
    data << clearData(sheet.row(5)[1])
    data << clearData(sheet.row(5)[2])
    data << clearData(sheet.row(5)[4])
    data << clearData(sheet.row(5)[6])
    data << clearData(sheet.row(5)[7])

    volumen = 'N/A'
    row_number = 0
    sheet.each do |row|
      row_number = row.idx
      if row[column].is_a?(String) and (row[column].upcase.include?('V TOTAL') or row[column].upcase.include?('VOLUMEN TOTAL'))
        volumen = getValue(row[column + 1])
        break
      end
    end
    data << getValue(sheet.row(row_number + 1)[1])
    data << getValue(sheet.row(row_number + 3)[1])
    data << getValue(sheet.row(row_number + 4)[1])
    data << getValue(sheet.row(row_number + 5)[1])

    data << volumen
    data << getValue(sheet.row(row_number + 1)[column + 1])

    data << xls

    csv << data
  end
  puts "Total de archivos: #{total_count}"
end
