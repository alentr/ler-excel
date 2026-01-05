package org.example.mapper;

import org.apache.poi.ss.usermodel.Row;
import org.example.model.Person;
import org.example.reader.RowMapper;
import org.example.util.CellValueExtractor;

/**
 * Implementação de mapeador de linha para converter linhas do Excel em objetos Person.
 *
 * Esta classe implementa a interface RowMapper para fornecer lógica de
 * mapeamento específica para o modelo Person. Ela extrai valores das células
 * do Excel e cria objetos Person.
 *
 * Estrutura esperada do Excel:
 * - Coluna A (índice 0): Nome (String)
 * - Coluna B (índice 1): Idade (Integer)
 *
 * Como criar mapeadores para diferentes modelos:
 * 1. Crie uma nova classe implementando RowMapper<SeuModelo>
 * 2. No método mapRow(), extraia valores das células usando CellValueExtractor
 * 3. Crie e retorne seu objeto de modelo
 *
 * Exemplo para um modelo Product:
 * <pre>
 * public class ProductRowMapper implements RowMapper<Product> {
 *     public Product mapRow(Row row) {
 *         Product product = new Product();
 *         product.setCode(CellValueExtractor.getStringValue(row.getCell(0)));
 *         product.setName(CellValueExtractor.getStringValue(row.getCell(1)));
 *         product.setPrice(CellValueExtractor.getDoubleValue(row.getCell(2)));
 *         return product;
 *     }
 * }
 * </pre>
 *
 * @author Seu Nome
 * @version 1.0
 */
public class PersonRowMapper implements RowMapper<Person> {

    /**
     * Índice da coluna para o campo Nome.
     * No Excel, a coluna A tem índice 0.
     */
    private static final int NAME_COLUMN_INDEX = 0;

    /**
     * Índice da coluna para o campo Idade.
     * No Excel, a coluna B tem índice 1.
     */
    private static final int AGE_COLUMN_INDEX = 1;

    /**
     * Mapeia uma linha do Excel para um objeto Person.
     *
     * Este método extrai o nome e a idade das células da linha
     * e cria um novo objeto Person com esses valores.
     *
     * O utilitário CellValueExtractor é usado para extrair valores
     * de forma segura, lidando com diferentes tipos de células e valores nulos.
     *
     * @param row A linha do Excel a ser mapeada
     * @return Um objeto Person preenchido com os dados da linha
     */
    @Override
    public Person mapRow(Row row) {
        // Cria um novo objeto Person para armazenar os dados extraídos
        Person person = new Person();

        // Extrai o nome da coluna A (índice 0)
        // CellValueExtractor lida com diferentes tipos de células (String, Numeric, etc.)
        // e retorna null para células vazias
        String name = CellValueExtractor.getStringValue(row.getCell(NAME_COLUMN_INDEX));
        person.setName(name);

        // Extrai a idade da coluna B (índice 1)
        // CellValueExtractor.getIntegerValue() converte células numéricas para Integer
        // e lida com células de texto que contêm números
        Integer age = CellValueExtractor.getIntegerValue(row.getCell(AGE_COLUMN_INDEX));
        person.setAge(age);

        // Retorna o objeto Person preenchido
        return person;
    }
}
