require 'csv'
require 'xml'
require 'rspreadsheet'

data_home = '/media/WD2T/Stockage/Hubic/Gestion/Compta/2019/formation/'
fec_file = '480751767FEC20191231.txt'
ods_file='tmp.ods'

class SheetDesc
    attr_accessor :title, :data
    def initialize(title)
        @title = title
        @data = $all.filter { |line| line['CompteNum'].start_with?(title == 'CA' ? '706' : title) }
    end

    def ca?()
        @title == 'CA'
    end

    def SheetDesc.create(title) 
        desc = SheetDesc.new(title)
        sheet = $workbook.create_worksheet(desc.title)
        header_row = sheet.row(1)
        header_row.cellvalues = ['EcritureNum', 'EcritureDate', 'EcritureNum', 'EcritureLib', 'Debit', 'Credit', 'Type', 'Origine']
        
        row = header_row
        desc.data.each do |line|
            row = sheet.row(row.rowi + 1)
            row.cellvalues = [
                line['EcritureNum'], 
                line['EcritureDate'], 
                line['EcritureNum'], 
                line['EcritureLib'], 
                line['Debit'].gsub(',', '.').to_f, 
                line['Credit'].gsub(',', '.').to_f
            ]
        end

        desc
    end
end

def rowSynthese(row, sheetDesc) 
    maxRows = sheetDesc.data.size + 1
    row[1] = "#{sheetDesc.title}"
    col = sheetDesc.ca?() ? 'F' : 'E'
    row.cell(2).formula = "=SUM($#{sheetDesc.title}.#{col}$2:$#{sheetDesc.title}.#{col}#{maxRows})"
    formule_formation = "=SUMIF($#{sheetDesc.title}.G$2:$#{sheetDesc.title}.G#{maxRows};\"F\";$#{sheetDesc.title}.#{col}$2:$#{sheetDesc.title}.#{col}#{maxRows})"
    formule_mixte = "SUMIF($#{sheetDesc.title}.G$2:$#{sheetDesc.title}.G#{maxRows};\"M\";$#{sheetDesc.title}.#{col}$2:$#{sheetDesc.title}.#{col}#{maxRows})*$D$2"
    row.cell(3).formula = sheetDesc.ca? ? formule_formation : "#{formule_formation} + #{formule_mixte}"
end

def synthese(sheets) 
    sheet = $workbook.create_worksheet('Synthese')

    header_row = sheet.row(1)
    header_row.cellvalues = ['', 'Total', 'Formation']

    index = 1
    sheets.each do |desc|
        index += 1
        rowSynthese(sheet.row(index), desc)
    end
    sheet.row(2).cell(4).formula = '=C2/B2'

    maxRows = sheets[0].data.size + 1
    index += 3
    origins = ['Entreprises', 'Formation Pro', 'Pouvoirs public', 'Particuliers', 'Sous-traitance', 'Autres (étranger,...)']
    origins.each_with_index do |label, key|
        sheet.row(index + key).cell(1).value = "A#{key}"
        sheet.row(index + key).cell(2).value = label
        sheet.row(index + key).cell(3).formula = "=SUMIF($CA.H$2:H#{maxRows};$A#{index + key};F$2:F#{maxRows})"
    end
end

$all = []
CSV.foreach(data_home + fec_file, headers: true, col_sep: '|') do |row|
    $all.push(row)
end

$workbook = Rspreadsheet.new

sheets = [
    SheetDesc.create('CA'),
    SheetDesc.create('60'),
    SheetDesc.create('61'),
    SheetDesc.create('62'),
    SheetDesc.create('63'),
    SheetDesc.create('64'),
    SheetDesc.create('65'),
    SheetDesc.create('68'),
    SheetDesc.create('69')
]
synthese(sheets)

$workbook.create_worksheet('Cours')

$workbook.save(data_home + ods_file)
