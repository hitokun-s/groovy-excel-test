import spock.lang.Specification


/**
 * @author Hitoshi Wada
 */
class ReaderSpec extends Specification {

    def "HSSF formattest"(){
        given:
            def reader = new ExcelReader(new File(getClass().getClassLoader().getResource("HSSF.xls").getFile()))
        when:
            reader.exec()
        then:
            true
    }
}