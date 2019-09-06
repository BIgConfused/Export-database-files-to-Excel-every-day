package utils.util;

import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

@Component
public class ScheRun {
    //每分钟执行
    @Scheduled(cron = "0 * * * * ? ")
    public void sRun(){
        POIDbToExcel.bakDBToExcel("xlsx","C:\\Users\\11783\\Desktop\\");
    }
}
