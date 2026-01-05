package org.example.reader;

import org.apache.poi.ss.usermodel.Row;

/**
 * Interface funcional para mapear linhas do Excel para objetos Java.
 *
 * Esta interface define um contrato para converter uma linha do Excel
 * em um tipo específico de objeto Java. Ela segue o padrão Strategy,
 * permitindo diferentes implementações de mapeamento para diferentes tipos de objetos.
 *
 * Ao implementar esta interface, você pode criar mapeadores personalizados
 * para qualquer tipo de objeto, tornando o leitor de Excel altamente flexível
 * e reutilizável em diferentes estruturas de arquivo.
 *
 * Exemplo de uso:
 * <pre>
 * RowMapper<Person> personMapper = row -> {
 *     Person person = new Person();
 *     person.setName(CellValueExtractor.getStringValue(row.getCell(0)));
 *     person.setAge(CellValueExtractor.getIntegerValue(row.getCell(1)));
 *     return person;
 * };
 * </pre>
 *
 * @param <T> O tipo de objeto para o qual a linha será mapeada
 * @author Seu Nome
 * @version 1.0
 */
@FunctionalInterface
public interface RowMapper<T> {

    /**
     * Mapeia uma linha do Excel para um objeto do tipo T.
     *
     * Este método é chamado para cada linha de dados no arquivo Excel
     * (tipicamente excluindo a linha de cabeçalho). A implementação
     * deve extrair valores das células da linha e popular um objeto
     * do tipo alvo.
     *
     * @param row A linha do Excel a ser mapeada
     * @return Um objeto do tipo T preenchido com os dados da linha
     */
    T mapRow(Row row);
}
