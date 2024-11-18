# ivy-vba-projects
 Development for VBA projects. 

Project Plan
1. Establish Database Connection
Status: Completed
Description: Create a UserForm to collect user credentials (user ID and password) and establish a connection to the Oracle database using the DatabaseManager class.
2. User Authentication
Status: In Progress
Description: Implement the UserManager class to validate user credentials against a predefined list or database.
3. Dynamic Data Population
Status: Planned
Description: Use the DataHandler class to dynamically populate Excel sheets with data fetched from the Oracle database. Ensure headers and data are correctly formatted.
4. Configuration Management
Status: Planned
Description: Implement the ConfigManager class to load and manage configurations from a codebook. Apply configurations dynamically based on the SQL query.
5. Error Handling
Status: Planned
Description: Centralize error handling using the ErrorHandler module to ensure consistent and robust error management across the project.
6. Transaction Management
Status: Planned
Description: Ensure that database transactions are only committed if all operations are successful. Implement transaction management in the DatabaseManager class.
7. User Interface Enhancements
Status: Planned
Description: Improve the UserForm and other user interface elements to enhance user experience and usability.
8. Documentation and Testing
Status: Planned
Description: Document the code and project thoroughly. Implement unit tests and perform comprehensive testing to ensure reliability and correctness.
Summary
This project aims to develop a robust and user-friendly VBA application that interacts with an Oracle database. The key components include user authentication, dynamic data population, configuration management, error handling, and transaction management. The project will be developed iteratively, with thorough documentation and testing at each stage.

Strict Separation of Concerns:

Keep transaction logic in DatabaseManager.
Handle dynamic configs in ConfigManager.
Let DataHandler focus purely on data population with minimal reliance on other modules.
Safe Transactions:

Ensure database operations use transactions that cleanly rollback in the event of an error.
Avoid unnecessary state coupling or redundant handling across modules.
Dynamic and Flexible Configurations:

Allow configurations to be fetched dynamically from Oracle and local Excel worksheets.
Avoid hardcoding values or introducing unnecessary manual dependencies.
Maintainability and Testability:

Ensure that unit tests validate functionality with mock setups while adhering to real-world behavior.
Use clean and clear test scenarios that reflect dynamic data usage.
Let’s correct the implementation while staying faithful to these established ideas.