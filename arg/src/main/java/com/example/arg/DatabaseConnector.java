package com.example.arg;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DatabaseConnector {
    private static Connection connection;

    public void connect() throws SQLException {
        String url = "jdbc:mysql://123.456.7.891:2345/datastore"; //update with real URL
        String user = "";
        String password = "";
        connection = DriverManager.getConnection(url, user, password);
        connection.setReadOnly(true);
    }

    public static Connection getConnection() {
        return connection;
    }

    public void closeConnection() throws SQLException {
        if (connection != null && !connection.isClosed()) {
            connection.close();
        }
    }

    // Method to check connection status
    public static boolean isConnectionOpen() {
        try {
            return connection != null && !connection.isClosed();
        } catch (SQLException e) {
            // Log exception (optional)
            System.err.println("Error checking connection status: " + e.getMessage());
            return false; // Assume it's closed in case of an error
        }
    }
}
