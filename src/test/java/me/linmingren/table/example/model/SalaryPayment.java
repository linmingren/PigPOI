package me.linmingren.table.example.model;

import lombok.Data;

@Data
public class SalaryPayment {
    private String userName;
    private int baseSalary;
    private int fullAttendanceBonus;
    private int mealSupplement;
    private int transportationAllowance;
    private int sickLeave;
    private int personalLeave;

    public SalaryPayment(String userName, int baseSalary, int fullAttendanceBonus, int mealSupplement,
                         int transportationAllowance,  int sickLeave, int personalLeave) {
        this.userName = userName;
        this.baseSalary = baseSalary;
        this.fullAttendanceBonus = fullAttendanceBonus;
        this.mealSupplement = mealSupplement;
        this.transportationAllowance = transportationAllowance;
        this.sickLeave = sickLeave;
        this.personalLeave = personalLeave;
    }

    public int getActualPay() {
        return this.baseSalary + this.fullAttendanceBonus + this.mealSupplement + this.transportationAllowance
                - sickLeave - personalLeave;
    }
}
