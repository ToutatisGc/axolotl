package cn.toutatis.xvoid.axolotl.excel.entities.writer;

import cn.hutool.core.util.RandomUtil;
import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictMapping;
import cn.toutatis.xvoid.toolkit.constant.Time;
import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDateTime;

@Data
public class StockEntity {

    private String code;

    private String intro;

    @AxolotlDictMapping(staticDict = {"1","是","0","否"})
    private String st;

    private Double pts;

    private LocalDateTime localDateTime = LocalDateTime.now();

    private String localDateTimeStr;

    private Double closingPrice = RandomUtil.randomDouble(0,100);

    private Double priceLimit = RandomUtil.randomDouble(0,100);

    private BigDecimal totalValue = RandomUtil.randomBigDecimal(BigDecimal.ZERO,new BigDecimal(100));

    private BigDecimal circulationMarketValue = RandomUtil.randomBigDecimal(BigDecimal.ZERO,new BigDecimal(100));

    public StockEntity() {
        localDateTimeStr = Time.regexTime(Time.SIMPLE_DATE_FORMAT_REGEX, localDateTime);
    }
}
