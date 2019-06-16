import parser.Schedule;

public class Main {

    public static void main(String[] arguments) {
        System.out.println(new Schedule(arguments[0]).export());
    }

}
