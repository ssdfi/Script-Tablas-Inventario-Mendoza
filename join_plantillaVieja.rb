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

  csv << ['concatenar', 'esp_m', 'orient', 'n_parcela', 'n_gps',
    'n_planta', 'dap_pr', 'volumen_total', 'altura_promedio', 'archivo']

  Dir.glob(files_dir) do |xls|
    next if File.extname(xls) != '.xls'
    total_count += 1
    book = Spreadsheet.open xls
    sheet = book.worksheet 0
    # data = File.basename(xls, '.xls').split(splitter).map!(&:strip)

    data = []

    # Leer datos de la cabecera
    data << clearData(sheet.row(1)[5])
    data << clearData(sheet.row(1)[6])
    data << clearData(sheet.row(1)[7])
    data << clearData(sheet.row(3)[1])
    data << clearData(sheet.row(3)[3])

    volumen = 'N/A'
    row_number = 0
    sheet.each do |row|
      row_number = row.idx
      if row[column].is_a?(String) and (row[column].upcase.include?('V TOTAL') or row[column].upcase.include?('VOLUMEN TOTAL'))
        volumen = getValue(row[column + 1])
        break
      end
    end
    data << clearData(sheet.row(row_number + 1)[1])
    data << clearData(sheet.row(row_number + 1)[5])


    data << volumen
    #Leo altura promedio que puede ser que esté en dos posiciones, si es nil, mando a la siguiente
    @alt_prom= sheet.row(row_number + 1)[column + 1].nil?
    if @alt_prom == false
      data << getValue(sheet.row(row_number + 1)[column + 1])
    else
      data << getValue(sheet.row(row_number + 2)[column + 1])
    end

    data << xls

    csv << data
  end
  puts "Total de archivos: #{total_count}"
end
