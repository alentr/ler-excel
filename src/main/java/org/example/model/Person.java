package org.example.model;

/**
 * Classe de modelo representando uma pessoa com informações básicas.
 *
 * Esta classe é um POJO (Plain Old Java Object) simples que armazena
 * os dados lidos de um arquivo Excel. Ela contém:
 * - name: O nome da pessoa (String)
 * - age: A idade da pessoa (Integer)
 *
 * Esta classe pode ser facilmente estendida para incluir mais campos
 * conforme necessário para diferentes estruturas de arquivo Excel.
 *
 * @author Seu Nome
 * @version 1.0
 */
public class Person {

    /**
     * O nome completo da pessoa.
     * Corresponde à coluna A no arquivo Excel.
     */
    private String name;

    /**
     * A idade da pessoa em anos.
     * Corresponde à coluna B no arquivo Excel.
     */
    private Integer age;

    /**
     * Construtor padrão.
     * Cria um objeto Person vazio.
     */
    public Person() {
    }

    /**
     * Construtor com parâmetros.
     * Cria um objeto Person com o nome e idade especificados.
     *
     * @param name O nome da pessoa
     * @param age  A idade da pessoa
     */
    public Person(String name, Integer age) {
        this.name = name;
        this.age = age;
    }

    /**
     * Obtém o nome da pessoa.
     *
     * @return O nome da pessoa
     */
    public String getName() {
        return name;
    }

    /**
     * Define o nome da pessoa.
     *
     * @param name O nome a ser definido
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * Obtém a idade da pessoa.
     *
     * @return A idade da pessoa
     */
    public Integer getAge() {
        return age;
    }

    /**
     * Define a idade da pessoa.
     *
     * @param age A idade a ser definida
     */
    public void setAge(Integer age) {
        this.age = age;
    }

    /**
     * Retorna uma representação em string do objeto Person.
     * Útil para debugging e logging.
     *
     * @return Uma string formatada contendo as informações da pessoa
     */
    @Override
    public String toString() {
        return "Person{" +
                "name='" + name + '\'' +
                ", age=" + age +
                '}';
    }
}
