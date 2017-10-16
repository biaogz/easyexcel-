package function.model;

import java.math.BigDecimal;
import java.util.Date;

import com.alibaba.excel.annotation.ExcelColumnNum;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * Created by jipengfei on 17/3/28.
 *
 * @author jipengfei
 * @date 2017/03/28
 */
public class LoanInfo extends BaseRowModel implements Comparable<LoanInfo> {
    @ExcelColumnNum(0)
    private String bankLoanId; // ���зſ���
    @ExcelColumnNum(1)
    private Long customerId;//���������
    @ExcelColumnNum(value = 2, format = "yyyy/MM/dd")
    private Date loanDate;// ���зſ�����
    @ExcelColumnNum(3)
    private BigDecimal quota; // ���зſ���
    @ExcelColumnNum(4)
    private String bankInterestRate;// ��������
    @ExcelColumnNum(5)
    private Integer loanTerm; // ���н������
    @ExcelColumnNum(value = 6, format = "yyyy/MM/dd")
    private Date loanEndDate;// ������
    @ExcelColumnNum(7)
    private BigDecimal interestPerMonth;// ÿ��Ӧ����Ϣ

    //    public String getLoanName() {
    //        return loanName;
    //    }
    //
    //    public void setLoanName(String loanName) {
    //        this.loanName = loanName;
    //    }

    public Date getLoanDate() {
        return loanDate;
    }

    public void setLoanDate(Date loanDate) {
        this.loanDate = loanDate;
    }

    public BigDecimal getQuota() {
        return quota;
    }

    public void setQuota(BigDecimal quota) {
        this.quota = quota;
    }

    public String getBankInterestRate() {
        return bankInterestRate;
    }

    public void setBankInterestRate(String bankInterestRate) {
        this.bankInterestRate = bankInterestRate;
    }

    public Integer getLoanTerm() {
        return loanTerm;
    }

    public void setLoanTerm(Integer loanTerm) {
        this.loanTerm = loanTerm;
    }

    public Date getLoanEndDate() {
        return loanEndDate;
    }

    public void setLoanEndDate(Date loanEndDate) {
        this.loanEndDate = loanEndDate;
    }

    public BigDecimal getInterestPerMonth() {
        return interestPerMonth;
    }

    public void setInterestPerMonth(BigDecimal interestPerMonth) {
        this.interestPerMonth = interestPerMonth;
    }

    public String getBankLoanId() {
        return bankLoanId;
    }

    public void setBankLoanId(String bankLoanId) {
        this.bankLoanId = bankLoanId;
    }

    public Long getCustomerId() {
        return customerId;
    }

    public void setCustomerId(Long customerId) {
        this.customerId = customerId;
    }

    @Override
    public String toString() {
        return "ExcelLoanInfo{" +
            "bankLoanId='" + bankLoanId + '\'' +
            ", customerId='" + customerId + '\'' +
            ", loanDate=" + loanDate +
            ", quota=" + quota +
            ", bankInterestRate=" + bankInterestRate +
            ", loanTerm=" + loanTerm +
            ", loanEndDate=" + loanEndDate +
            ", interestPerMonth=" + interestPerMonth +
            '}';
    }

    public int compareTo(LoanInfo info) {
        boolean before = this.getLoanDate().before(info.getLoanDate());
        return before ? 1 : 0;
    }
}