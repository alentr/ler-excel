package org.example.reader;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Classe genérica para leitura de arquivos Excel.
 *
 * Esta classe fornece métodos para ler arquivos Excel (.xls e .xlsx)
 * e converter seus conteúdos em objetos Java usando um RowMapper.
 *
 * O leitor foi projetado para ser flexível e reutilizável:
 * - Funciona com arquivos .xls (Excel 97-2003) e .xlsx (Excel 2007+)
 * - Usa tipos genéricos para suportar qualquer tipo de objeto
 * - Aceita um RowMapper para personalizar como as linhas são convertidas em objetos
 * - Suporta pular linhas de cabeçalho
 * - Trata erros de forma elegante
 *
 * Padrões de Design utilizados:
 * - Padrão Strategy: RowMapper permite diferentes estratégias de mapeamento
 * - Template Method: Lógica de leitura comum com mapeamento personalizável
 *
 * Exemplo de uso:
 * <pre>
 * ExcelReader reader = new ExcelReader();
 * List<Person> people = reader.readFile("caminho/para/arquivo.xlsx", new PersonRowMapper());
 * </pre>
 *
 * @author Seu Nome
 * @version 1.0
 */
public class ExcelReader {

    /**
     * Número padrão de linhas de cabeçalho a serem puladas.
     * A maioria dos arquivos Excel tem uma linha de cabeçalho com nomes de colunas.
     */
    private static final int DEFAULT_HEADER_ROWS = 1;

    /**
     * Lê um arquivo Excel e converte suas linhas em uma lista de objetos.
     *
     * Este método lê a primeira planilha do arquivo Excel, pula a(s) linha(s)
     * de cabeçalho, e usa o RowMapper fornecido para converter cada linha
     * de dados em um objeto do tipo T.
     *
     * @param filePath  O caminho para o arquivo Excel
     * @param rowMapper O mapeador para converter linhas em objetos
     * @param <T>       O tipo de objetos a serem criados
     * @return Uma lista de objetos criados a partir das linhas do Excel
     * @throws IOException Se o arquivo não puder ser lido
     */
    public <T> List<T> readFile(String filePath, RowMapper<T> rowMapper) throws IOException {
        return readFile(filePath, rowMapper, DEFAULT_HEADER_ROWS);
    }

    /**
     * Lê um arquivo Excel com um número personalizado de linhas de cabeçalho a serem puladas.
     *
     * Este método fornece mais controle sobre como o arquivo é lido,
     * permitindo especificar quantas linhas devem ser puladas antes
     * de começar a ler os dados.
     *
     * @param filePath       O caminho para o arquivo Excel
     * @param rowMapper      O mapeador para converter linhas em objetos
     * @param headerRowCount Número de linhas de cabeçalho a serem puladas (0 se não houver cabeçalhos)
     * @param <T>            O tipo de objetos a serem criados
     * @return Uma lista de objetos criados a partir das linhas do Excel
     * @throws IOException Se o arquivo não puder ser lido
     */
    public <T> List<T> readFile(String filePath, RowMapper<T> rowMapper, int headerRowCount) throws IOException {
        return readFile(filePath, rowMapper, headerRowCount, 0);
    }

    /**
     * Lê uma planilha específica de um arquivo Excel.
     *
     * Este é o método mais flexível, permitindo especificar:
     * - Qual planilha ler (por índice)
     * - Quantas linhas de cabeçalho pular
     * - A estratégia de mapeamento
     *
     * @param filePath       O caminho para o arquivo Excel
     * @param rowMapper      O mapeador para converter linhas em objetos
     * @param headerRowCount Número de linhas de cabeçalho a serem puladas
     * @param sheetIndex     Índice da planilha a ser lida (baseado em 0)
     * @param <T>            O tipo de objetos a serem criados
     * @return Uma lista de objetos criados a partir das linhas do Excel
     * @throws IOException Se o arquivo não puder ser lido
     */
    public <T> List<T> readFile(String filePath, RowMapper<T> rowMapper, int headerRowCount, int sheetIndex)
            throws IOException {

        // Lista para armazenar os objetos mapeados
        List<T> result = new ArrayList<>();

        // Usa try-with-resources para garantir que os streams sejam fechados corretamente
        // FileInputStream abre o arquivo para leitura
        // WorkbookFactory.create() detecta automaticamente o tipo de arquivo (.xls ou .xlsx)
        try (InputStream inputStream = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(inputStream)) {

            // Obtém a planilha especificada do workbook
            // Planilhas são indexadas a partir de 0, então a primeira planilha está no índice 0
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            // Registra informações sobre a planilha sendo processada
            System.out.println("Lendo planilha: " + sheet.getSheetName());
            System.out.println("Total de linhas (incluindo cabeçalho): " + (sheet.getLastRowNum() + 1));

            // Obtém um iterador sobre todas as linhas da planilha
            Iterator<Row> rowIterator = sheet.iterator();

            // Pula as linhas de cabeçalho
            // Linhas de cabeçalho tipicamente contêm nomes de colunas, não dados
            int skippedRows = 0;
            while (rowIterator.hasNext() && skippedRows < headerRowCount) {
                rowIterator.next(); // Pula esta linha
                skippedRows++;
            }

            // Processa cada linha de dados
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Pula linhas completamente vazias
                // Uma linha é considerada vazia se não tiver células ou todas as células estiverem em branco
                if (isRowEmpty(row)) {
                    continue;
                }

                // Usa o RowMapper para converter a linha em um objeto
                // O mapeador encapsula a lógica de extração de valores das células
                T mappedObject = rowMapper.mapRow(row);

                // Adiciona o objeto mapeado à lista de resultados
                if (mappedObject != null) {
                    result.add(mappedObject);
                }
            }

            System.out.println("Lidos com sucesso " + result.size() + " registros.");
        }

        return result;
    }

    /**
     * Verifica se uma linha está vazia (não tem dados).
     *
     * Uma linha é considerada vazia se:
     * - For null
     * - Não tiver células
     * - Todas as células estiverem em branco ou contiverem apenas espaços em branco
     *
     * @param row A linha a ser verificada
     * @return true se a linha estiver vazia, false caso contrário
     */
    private boolean isRowEmpty(Row row) {
        // Linha null está vazia
        if (row == null) {
            return true;
        }

        // Verifica cada célula da linha
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            var cell = row.getCell(i);
            if (cell != null) {
                // Verifica se a célula tem algum conteúdo
                switch (cell.getCellType()) {
                    case STRING:
                        if (!cell.getStringCellValue().trim().isEmpty()) {
                            return false; // Linha tem conteúdo
                        }
                        break;
                    case NUMERIC:
                    case BOOLEAN:
                    case FORMULA:
                        return false; // Linha tem conteúdo
                    default:
                        break;
                }
            }
        }

        // Todas as células estão vazias
        return true;
    }

    /**
     * Obtém o número de planilhas em um arquivo Excel.
     *
     * Útil quando você precisa iterar sobre todas as planilhas de um workbook.
     *
     * @param filePath O caminho para o arquivo Excel
     * @return O número de planilhas no workbook
     * @throws IOException Se o arquivo não puder ser lido
     */
    public int getSheetCount(String filePath) throws IOException {
        try (InputStream inputStream = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            return workbook.getNumberOfSheets();
        }
    }

    /**
     * Obtém os nomes de todas as planilhas em um arquivo Excel.
     *
     * Útil para exibir opções de planilhas aos usuários ou para logging.
     *
     * @param filePath O caminho para o arquivo Excel
     * @return Uma lista com os nomes das planilhas
     * @throws IOException Se o arquivo não puder ser lido
     */
    public List<String> getSheetNames(String filePath) throws IOException {
        List<String> sheetNames = new ArrayList<>();

        try (InputStream inputStream = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(inputStream)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
        }

        return sheetNames;
    }
}

