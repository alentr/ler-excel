package org.example.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.util.Date;

/**
 * Classe utilitária para extração de valores de células do Excel.
 *
 * Esta classe fornece métodos estáticos para extrair de forma segura diferentes tipos
 * de valores de células do Excel, lidando com vários tipos de células e casos extremos.
 *
 * Células do Excel podem conter diferentes tipos de dados:
 * - STRING: Valores de texto
 * - NUMERIC: Números (incluindo datas, que são armazenadas como números no Excel)
 * - BOOLEAN: Valores Verdadeiro/Falso
 * - FORMULA: Valores calculados (extraímos o resultado em cache)
 * - BLANK: Células vazias
 * - ERROR: Células com erros
 *
 * Este utilitário lida com todos esses casos de forma elegante, fornecendo métodos
 * de extração null-safe que podem ser facilmente usados em toda a aplicação.
 *
 * @author Seu Nome
 * @version 1.0
 */
public class CellValueExtractor {

    /**
     * Construtor privado para prevenir instanciação.
     * Esta é uma classe utilitária com apenas métodos estáticos.
     */
    private CellValueExtractor() {
        // Classe utilitária - não instanciar
    }

    /**
     * Extrai um valor String de uma célula do Excel.
     *
     * Este método lida com diferentes tipos de células:
     * - STRING: Retorna o valor string diretamente
     * - NUMERIC: Converte o número para string (trata datas especialmente)
     * - BOOLEAN: Converte para "true" ou "false"
     * - FORMULA: Avalia e retorna o resultado como string
     * - BLANK/null: Retorna null
     *
     * @param cell A célula do Excel da qual extrair o valor
     * @return O valor da célula como String, ou null se a célula estiver vazia
     */
    public static String getStringValue(Cell cell) {
        // Retorna null para células vazias
        if (cell == null) {
            return null;
        }

        // Lida com diferentes tipos de células
        switch (cell.getCellType()) {
            case STRING:
                // Valor string direto
                return cell.getStringCellValue();

            case NUMERIC:
                // Verifica se é uma data (datas são armazenadas como números no Excel)
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue().toString();
                }
                // Converte número para string, removendo casas decimais desnecessárias
                double numValue = cell.getNumericCellValue();
                if (numValue == Math.floor(numValue)) {
                    // É um número inteiro, retorna sem casas decimais
                    return String.valueOf((long) numValue);
                }
                return String.valueOf(numValue);

            case BOOLEAN:
                // Converte booleano para string
                return String.valueOf(cell.getBooleanCellValue());

            case FORMULA:
                // Para fórmulas, tenta obter o resultado em cache
                try {
                    return cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    // Resultado da fórmula é numérico
                    return String.valueOf(cell.getNumericCellValue());
                }

            case BLANK:
                // Célula vazia
                return null;

            case ERROR:
                // Célula contém um erro
                return "ERRO";

            default:
                return null;
        }
    }

    /**
     * Extrai um valor Integer de uma célula do Excel.
     *
     * Este método converte células numéricas para valores Integer.
     * Também lida com células de texto que contêm valores numéricos.
     *
     * @param cell A célula do Excel da qual extrair o valor
     * @return O valor da célula como Integer, ou null se a célula estiver vazia ou não for numérica
     */
    public static Integer getIntegerValue(Cell cell) {
        // Retorna null para células vazias
        if (cell == null) {
            return null;
        }

        // Lida com diferentes tipos de células
        switch (cell.getCellType()) {
            case NUMERIC:
                // Converte double para integer (truncando casas decimais)
                return (int) cell.getNumericCellValue();

            case STRING:
                // Tenta fazer parse da string como integer
                try {
                    String stringValue = cell.getStringCellValue().trim();
                    // Lida com strings decimais fazendo parse como double primeiro
                    return (int) Double.parseDouble(stringValue);
                } catch (NumberFormatException e) {
                    // String não é um número válido
                    return null;
                }

            case FORMULA:
                // Para fórmulas, tenta obter o resultado numérico
                try {
                    return (int) cell.getNumericCellValue();
                } catch (IllegalStateException e) {
                    return null;
                }

            default:
                return null;
        }
    }

    /**
     * Extrai um valor Double de uma célula do Excel.
     *
     * Este método converte células numéricas para valores Double.
     * Útil para células contendo números decimais.
     *
     * @param cell A célula do Excel da qual extrair o valor
     * @return O valor da célula como Double, ou null se a célula estiver vazia ou não for numérica
     */
    public static Double getDoubleValue(Cell cell) {
        // Retorna null para células vazias
        if (cell == null) {
            return null;
        }

        // Lida com diferentes tipos de células
        switch (cell.getCellType()) {
            case NUMERIC:
                return cell.getNumericCellValue();

            case STRING:
                // Tenta fazer parse da string como double
                try {
                    return Double.parseDouble(cell.getStringCellValue().trim());
                } catch (NumberFormatException e) {
                    return null;
                }

            case FORMULA:
                try {
                    return cell.getNumericCellValue();
                } catch (IllegalStateException e) {
                    return null;
                }

            default:
                return null;
        }
    }

    /**
     * Extrai um valor Boolean de uma célula do Excel.
     *
     * Este método lida com células booleanas e também interpreta
     * representações string comuns de valores booleanos.
     *
     * @param cell A célula do Excel da qual extrair o valor
     * @return O valor da célula como Boolean, ou null se a célula estiver vazia ou não for booleana
     */
    public static Boolean getBooleanValue(Cell cell) {
        // Retorna null para células vazias
        if (cell == null) {
            return null;
        }

        // Lida com diferentes tipos de células
        switch (cell.getCellType()) {
            case BOOLEAN:
                return cell.getBooleanCellValue();

            case STRING:
                // Interpreta representações string comuns
                var value = cell.getStringCellValue().trim().toLowerCase();
                if ("true".equals(value) || "yes".equals(value) || "1".equals(value) || "sim".equals(value)) {
                    return true;
                } else if ("false".equals(value) || "no".equals(value) || "0".equals(value) || "não".equals(value)) {
                    return false;
                }
                return null;

            case NUMERIC:
                // 0 = false, qualquer outro valor = true
                return cell.getNumericCellValue() != 0;

            default:
                return null;
        }
    }

    /**
     * Extrai um valor Date de uma célula do Excel.
     *
     * O Excel armazena datas como valores numéricos (número de dias desde uma data base).
     * Este método converte corretamente tais células para objetos Java Date.
     *
     * @param cell A célula do Excel da qual extrair o valor
     * @return O valor da célula como Date, ou null se a célula estiver vazia ou não for uma data
     */
    public static Date getDateValue(Cell cell) {
        // Retorna null para células vazias
        if (cell == null) {
            return null;
        }

        // Lida com células numéricas que contêm datas
        if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            }
        }

        return null;
    }
}

