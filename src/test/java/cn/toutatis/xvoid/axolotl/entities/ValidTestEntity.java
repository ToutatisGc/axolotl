package cn.toutatis.xvoid.axolotl.entities;


import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import lombok.Data;

import javax.validation.constraints.Min;
import javax.validation.constraints.NotBlank;


@Data
public class ValidTestEntity {

    @NotBlank
    private String name;

    @Min(value = 1,message = "AAA")
    private int age;

}
