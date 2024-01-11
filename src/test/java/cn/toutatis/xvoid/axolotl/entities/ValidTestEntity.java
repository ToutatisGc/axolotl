package cn.toutatis.xvoid.axolotl.entities;


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
