package cn.sxw.zcs.helper.excel;

/**
 * @Author ZengCS
 * @Date 2023/3/28 16:07
 * @Copyright 四川生学教育科技有限公司
 * @Address 成都市天府软件园E3-3F
 */
public class StudentImportBean {
    /**
     * 年级
     */
    private String gradeName;
    /**
     * 班级
     */
    private Integer classNum;
    /**
     * 姓名
     */
    private String name;
    /**
     * 性别
     */
    private String gender;
    /**
     * 学籍号
     */
    private String code;
    /**
     * 民族
     */
    private String nation;

    public String getGradeName() {
        return gradeName;
    }

    public void setGradeName(String gradeName) {
        this.gradeName = gradeName;
    }

    public Integer getClassNum() {
        return classNum;
    }

    public void setClassNum(Integer classNum) {
        this.classNum = classNum;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getNation() {
        return nation;
    }

    public void setNation(String nation) {
        this.nation = nation;
    }

    @Override
    public String toString() {
        return "StudentImportBean{" +
                "gradeName='" + gradeName + classNum + "班" + '\'' +
                ", name='" + name + '\'' +
                ", gender='" + gender + '\'' +
                ", code='" + code + '\'' +
                ", nation='" + nation + '\'' +
                '}';
    }
}
