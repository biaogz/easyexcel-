package function.model;

import java.util.Date;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.FieldType;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * Created by jipengfei on 17/3/15.
 *
 * @author jipengfei
 * @date 2017/03/15
 */
public class AllDataTypeModel extends BaseRowModel {
    @ExcelProperty("name")
    private String name;

    @ExcelProperty(value = "startTime",format = "yyyy-MM-dd HH:mm:ss")
    private Date startTime;

    @ExcelProperty("endTime")
    private Date endTime;

    @ExcelProperty("url")
    private String url;

    @ExcelProperty(value = "times")
    private int times;

    @ExcelProperty("activityCode")
    private String activityCode;

    @ExcelProperty(value = "poplayerPageId")
    private long poplayerPageId;

    @ExcelProperty("aDouble")
    private Double aDouble;

    @ExcelProperty("aBoolean")
    private boolean aBoolean;

    @ExcelProperty("phoneNum")
    private String phoneNum;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getStartTime() {
        return startTime;
    }

    public void setStartTime(Date startTime) {
        this.startTime = startTime;
    }

    public Date getEndTime() {
        return endTime;
    }

    public void setEndTime(Date endTime) {
        this.endTime = endTime;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public int getTimes() {
        return times;
    }

    public void setTimes(int times) {
        this.times = times;
    }

    public String getActivityCode() {
        return activityCode;
    }

    public void setActivityCode(String activityCode) {
        this.activityCode = activityCode;
    }

    public long getPoplayerPageId() {
        return poplayerPageId;
    }

    public void setPoplayerPageId(long poplayerPageId) {
        this.poplayerPageId = poplayerPageId;
    }

    public Double getaDouble() {
        return aDouble;
    }

    public void setaDouble(Double aDouble) {
        this.aDouble = aDouble;
    }

    public Boolean getaBoolean() {
        return aBoolean;
    }

    public void setaBoolean(Boolean aBoolean) {
        this.aBoolean = aBoolean;
    }

    public boolean isaBoolean() {
        return aBoolean;
    }

    public void setaBoolean(boolean aBoolean) {
        this.aBoolean = aBoolean;
    }

    public String getPhoneNum() {
        return phoneNum;
    }

    public void setPhoneNum(String phoneNum) {
        this.phoneNum = phoneNum;
    }
}
