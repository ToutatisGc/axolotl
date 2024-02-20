package cn.toutatis.xvoid.axolotl.excel.entities;


import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import lombok.Data;


@Data
public class ValidTestEntity {

    @NotBlank
    private String name;

    @Min(value = 1,message = "AAA")
    private int age;

}
