package org.example;

import org.example.mapper.PersonRowMapper;
import org.example.reader.ExcelReader;

import java.io.IOException;

/**
 * Classe principal da aplicação demonstrando as capacidades de leitura de arquivos Excel.
 *
 * Esta aplicação mostra como:
 * 1. Ler um arquivo Excel (.xlsx ou .xls)
 * 2. Mapear linhas do Excel para objetos Java usando um RowMapper
 * 3. Processar os dados extraídos
 *
 * O código é estruturado para ser facilmente adaptável:
 * - Crie novas classes de modelo para diferentes estruturas de dados
 * - Crie novas implementações de RowMapper para diferentes layouts de Excel
 * - Use o mesmo ExcelReader para qualquer tipo de arquivo Excel
 *
 * Estrutura do Projeto:
 * - org.example.model: Classes de modelo de dados (POJOs)
 * - org.example.reader: Infraestrutura de leitura de Excel
 * - org.example.mapper: Implementações de RowMapper para modelos específicos
 * - org.example.util: Classes utilitárias para extração de valores de células
 *
 * @author Seu Nome
 * @version 1.0
 */
public class Main {

    /**
     * Ponto de entrada principal da aplicação.
     *
     * Este método demonstra a funcionalidade de leitura de Excel:
     * 1. Criando uma instância de ExcelReader
     * 2. Criando um PersonRowMapper para mapear linhas para objetos Person
     * 3. Lendo o arquivo Excel e obtendo uma lista de objetos Person
     * 4. Imprimindo os dados extraídos no console
     *
     * @param args Argumentos da linha de comando (não utilizados)
     */
    public static void main(String[] args) {
        // Caminho para o arquivo Excel
        // Você pode alterar isso para ler diferentes arquivos
        // Suporta formatos .xlsx (Excel 2007+) e .xls (Excel 97-2003)
        var filePath = "docs/exemplo.xlsx";

        // Cria a instância do leitor de Excel
        // Esta classe lida com a abertura do arquivo e iteração pelas linhas
        var excelReader = new ExcelReader();

        // Cria o mapeador de linhas para objetos Person
        // O mapeador define como extrair dados de cada linha
        // Para diferentes modelos, crie diferentes implementações de RowMapper
        var personMapper = new PersonRowMapper();

        try {
            // Exibe informações sobre o arquivo sendo lido
            System.out.println("=========================================");
            System.out.println("       Demo Leitor de Arquivo Excel");
            System.out.println("=========================================");
            System.out.println();
            System.out.println("Lendo arquivo: " + filePath);
            System.out.println();

            // Lê o arquivo Excel e converte as linhas para objetos Person
            // O método readFile:
            // 1. Abre o arquivo Excel
            // 2. Lê a primeira planilha
            // 3. Pula a linha de cabeçalho (primeira linha)
            // 4. Usa o mapeador para converter cada linha de dados em um objeto Person
            // 5. Retorna uma lista de todos os objetos Person
            var people = excelReader.readFile(filePath, personMapper);

            // Exibe os resultados
            System.out.println();
            System.out.println("=========================================");
            System.out.println("           Dados Extraídos");
            System.out.println("=========================================");
            System.out.println();

            // Verifica se foram encontrados dados
            if (people.isEmpty()) {
                System.out.println("Nenhum dado encontrado no arquivo Excel.");
            } else {
                // Imprime o cabeçalho da tabela
                System.out.printf("%-5s | %-20s | %-10s%n", "#", "Nome", "Idade");
                System.out.println("------+----------------------+------------");

                // Itera pela lista e imprime cada pessoa
                var index = 1;
                for (var person : people) {
                    // Formata a saída como uma tabela
                    System.out.printf("%-5d | %-20s | %-10s%n",
                            index,
                            person.getName() != null ? person.getName() : "N/A",
                            person.getAge() != null ? person.getAge() : "N/A");
                    index++;
                }

                // Imprime o resumo
                System.out.println();
                System.out.println("Total de registros: " + people.size());
            }

        } catch (IOException e) {
            // Trata erros de leitura de arquivo
            System.err.println("Erro ao ler arquivo Excel: " + e.getMessage());
            System.err.println();
            System.err.println("Possíveis causas:");
            System.err.println("- Arquivo não encontrado no caminho especificado");
            System.err.println("- Arquivo está aberto em outra aplicação");
            System.err.println("- Arquivo está corrompido ou não é um arquivo Excel válido");
            System.err.println();
            System.err.println("Por favor, verifique o caminho do arquivo: " + filePath);

            // Imprime o stack trace para debugging
            e.printStackTrace();
        }

        System.out.println();
        System.out.println("=========================================");
        System.out.println("         Programa finalizado");
        System.out.println("=========================================");
    }
}
