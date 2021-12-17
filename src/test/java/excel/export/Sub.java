package excel.export;

public class Sub extends Parent {

    @Override
    public String getName() {
        return name;
    }

    @Override
    public void setName(String name) {
        this.name = name;
        super.setName(this.name);
    }

    /**
     * 姓名
     */
    private String name;

    public Sub(String name) {
        super(name);
    }


}
