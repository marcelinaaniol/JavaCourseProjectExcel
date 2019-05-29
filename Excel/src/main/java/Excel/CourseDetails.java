package Excel;

class CourseDetails {

    public String TIME, SUBJECTDESCRIPTION, DAYOFWEEK, DATE;
    public int NUMBEROFHOURS;

    public CourseDetails(String TIME, String SUBJECTDESCRIPTION, String DAYOFWEEK, int NUMBEROFHOURS, String DATE)

    {
        this.TIME = TIME;
        this.SUBJECTDESCRIPTION = SUBJECTDESCRIPTION;
        this.DAYOFWEEK = DAYOFWEEK;
        this.NUMBEROFHOURS = NUMBEROFHOURS;
        this.DATE = DATE;
    }

}