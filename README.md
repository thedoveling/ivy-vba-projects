# ivy-vba-projects
 Development for VBA projects. 

Project Plan
Suggested Refactoring Using Pragmatic Programmer Principles
1. Single Responsibility Principle
Each class should focus on a single responsibility:

DatabaseManager: Connection management and query execution.
ConfigManager: Metadata retrieval and configuration management.
DataHandler: Data processing and worksheet population.

2. Clear Lifecycle Flow
Define a clear lifecycle:

Open database connection (DatabaseManager).
Fetch and cache metadata (ConfigManager).
Query data and populate the worksheet (DataHandler).

3. Explicit Dependencies
Pass dependencies explicitly (e.g., pass DatabaseManager to ConfigManager and DataHandler).

4. Avoid Redundant State
Remove duplicate checks for connection state across classes. Rely on a single point of truth (DatabaseManager).

