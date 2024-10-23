package com.tripillar.filehandling.text;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;

public class WriteTextFile {
    public static void main(String[] args) {
        try {
            FileWriter writer = new FileWriter("src/WriteTextFile.txt", true);
            BufferedWriter bufferedWriter = new BufferedWriter(writer);

            bufferedWriter.write("Hello ! This is Ayush Raj.");
            bufferedWriter.newLine();
            bufferedWriter.write("I am Writing in a Text file.");

            bufferedWriter.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}