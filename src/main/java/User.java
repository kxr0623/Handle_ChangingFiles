public class User {
    private String name;
    private String account;
    private String Department;
    private boolean gender;
    private String email;
    private String Mobile;
    private String Password;
    private String State;
    public static final String USER_STATE_VALIDE="Valid";
//    public User(String name,String account,String Department, boolean gender,String email){
//        this.name=name;
//        this.account=account;
//        this.Department=Department;
//        this.gender=gender;
//        this.email=email;
//    }

    public String getName() {
        return name;
    }
    public String getAccount(){
        return account;
    }

    public String getDept() {
        return Department;
    }

    public boolean isGender() {
        return gender;

    }

    public String getEmail() {
        return email;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setAccount(String account) {
        this.account = account;
    }

    public void setDept(String department) {
        Department = department;
    }

    public void setGender(boolean gender) {
        this.gender = gender;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public void setMobile(String mobile) {
        Mobile = mobile;
    }

    public void setPassword(String password) {
        Password = password;
    }

    public void setState(String state) {
        State = state;
    }

}
