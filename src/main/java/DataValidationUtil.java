/**
 * Created by Any on 2018/7/15.
 */
import org.apache.poi.hssf.usermodel.DVConstraint;

public class DataValidationUtil {

    public static DVConstraint getListDVConstraint(String[] strArray){
        DVConstraint constraint = DVConstraint.createExplicitListConstraint(strArray);
        return constraint;
    }

    public static DVConstraint getDateDVConstraintBetween(String beginDate,String endDate,String dateFormat){
        DVConstraint constraint = DVConstraint.createDateConstraint(DVConstraint.OperatorType.BETWEEN, beginDate,
                endDate, dateFormat);
        return constraint;
    }

    public static DVConstraint getIntDVConstraintBetween(String beginData,String endData){
        DVConstraint constraint = DVConstraint.createNumericConstraint(DVConstraint.ValidationType.INTEGER, DVConstraint.OperatorType.BETWEEN, beginData, endData);
        return constraint;
    }

    public static DVConstraint getDecimalDVConstraintBetween(String beginData,String endData){
        DVConstraint constraint = DVConstraint.createNumericConstraint(DVConstraint.ValidationType.DECIMAL, DVConstraint.OperatorType.BETWEEN, beginData, endData);
        return constraint;
    }
}
