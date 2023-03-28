package cn.sxw.zcs.helper.excel;

/**
 * @Author ZengCS
 * @Date 2023/3/28 16:42
 * @Copyright 四川生学教育科技有限公司
 * @Address 成都市天府软件园E3-3F
 */
public class CompleteState {
    /**
     * 已完成-正常(在信息采集里面，导出已完成)
     */
    public static final Integer NORMAL_COMPLETE = 0;

    /**
     * 未完成-正常(在信息采集里面，导出未完成)
     */
    public static final Integer NORMAL_UN_COMPLETE = 1;

    /**
     * 已完成-异常(不在信息采集里面，导出已完成)
     */
    public static final Integer ERROR_COMPLETE = 2;

    /**
     * 未完成-异常(不在信息采集里面，导出未完成)
     */
    public static final Integer ERROR_UN_COMPLETE = 3;
}
