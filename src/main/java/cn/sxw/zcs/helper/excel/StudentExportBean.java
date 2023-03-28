package cn.sxw.zcs.helper.excel;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author ZengCS
 * @Date 2023/3/28 16:07
 * @Copyright 四川生学教育科技有限公司
 * @Address 成都市天府软件园E3-3F
 */
public class StudentExportBean {
    /**
     * 学籍号
     */
    private String code;
    /**
     * 学校
     */
    private String school;
    /**
     * 年级/班级
     */
    private String gradeClass;
    /**
     * 用户ID
     */
    private String userId;

    /**
     * 姓名
     */
    private String name;
    /**
     * 性别
     */
    private String gender;
    /**
     * 家长联系电话
     */
    private String parentPhone;
    /**
     * 最后作答时间
     */
    private String lastAnswerTime;
    /**
     * 作答用时
     */
    private String useTime;
    /**
     * 状态：已完成、未完成
     */
    private Integer completeState;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getSchool() {
        return school;
    }

    public void setSchool(String school) {
        this.school = school;
    }

    public String getGradeClass() {
        return gradeClass;
    }

    public void setGradeClass(String gradeClass) {
        this.gradeClass = gradeClass;
    }

    public String getUserId() {
        return userId;
    }

    public void setUserId(String userId) {
        this.userId = userId;
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

    public String getParentPhone() {
        return parentPhone;
    }

    public void setParentPhone(String parentPhone) {
        this.parentPhone = parentPhone;
    }

    public String getLastAnswerTime() {
        return lastAnswerTime;
    }

    public void setLastAnswerTime(String lastAnswerTime) {
        this.lastAnswerTime = lastAnswerTime;
    }

    public String getUseTime() {
        return useTime;
    }

    public void setUseTime(String useTime) {
        this.useTime = useTime;
    }

    public Integer getCompleteState() {
        return completeState;
    }

    public void setCompleteState(Integer completeState) {
        this.completeState = completeState;
    }

    public List<String> getAttributes() {
        // {"学籍号", "学校", "年级", "用户Id", "姓名", "性别", "家长联系电话", "最后作答时间", "用时（分钟）"};
        List<String> list = new ArrayList<>();
        list.add(name);
        list.add(code);
        list.add(school);
        list.add(gradeClass);
        list.add(userId);
        list.add(gender);
        list.add(parentPhone);
        list.add(lastAnswerTime);
        list.add(useTime);
        return list;
    }

    @Override
    public String toString() {
        return "StudentExportBean{" +
                "code='" + code + '\'' +
                ", school='" + school + '\'' +
                ", gradeClass='" + gradeClass + '\'' +
                ", userId='" + userId + '\'' +
                ", name='" + name + '\'' +
                ", gender='" + gender + '\'' +
                ", parentPhone='" + parentPhone + '\'' +
                ", lastAnswerTime='" + lastAnswerTime + '\'' +
                ", useTime='" + useTime + '\'' +
                ", completeState=" + completeState +
                '}';
    }
}
