/*
 * Copyright (c) 2019 Nadav Tasher
 * https://github.com/NadavTasher/HandasaimScheduler
 * https://github.com/NadavTasher/HandasaimWeb
 */

import org.json.JSONObject;
import parser.Schedule;

import java.io.File;
import java.nio.file.Files;

public class Main {
    public static void main(String[] arguments) {
        JSONObject json = new JSONObject();
        try {
            if (arguments.length >= 1) {
                json = new Schedule(arguments[0]);
            }
            if(json.getJSONArray(Schedule.ERRORS).length() == 0) {
                if (arguments.length >= 2) {
                    try {
                        Files.write(new File(arguments[1]).toPath(), json.getBytes());
                    } catch (Exception ignored) {
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println(json);
    }
}
