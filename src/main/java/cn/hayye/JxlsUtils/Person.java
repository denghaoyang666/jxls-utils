package cn.hayye.JxlsUtils;

public class Person {
    private Long id;
    private String name;
    private String nick;
    private String sex;
    private Integer age;
    private String birth;

    public Person(Long id, String name,String nick, String sex, Integer age, String birth) {
        this.id = id;
        this.name = name;
        this.nick = nick;
        this.sex = sex;
        this.age = age;
        this.birth = birth;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getNick() {
        return nick;
    }

    public void setNick(String nick) {
        this.nick = nick;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getBirth() {
        return birth;
    }

    public void setBirth(String birth) {
        this.birth = birth;
    }
}
