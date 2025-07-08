# inventory-management
This is a Google Apps Script code for an **inventory management system** in Google Sheets, designed for Persian users.  

### **Key Features**:  
1. **Custom Menu**: Adds an "Inventory" menu with options to:  
   - **Update Product**: Modify existing product data (stock, returns, monthly production/sales).  
   - **Create New Product**: Add a new product with a unique ID.  

2. **Dropdown List**:  
   - A Persian-language dropdown in "Data-Form" for selecting fields like initial stock, returns, and monthly production/sales.  

3. **Logging System**:  
   - Tracks all changes in a "logs" sheet with:  
     - Operation type (update/create)  
     - Gregorian & Jalali (Persian) dates  
     - Product ID, name, and modified values  

4. **Error Handling**:  
   - Validates inputs (checks for duplicates, empty fields, etc.).  
   - Shows user-friendly alerts in Persian.  

### **Sheets Used**:  
- **Data-Form**: Input form for updates/new entries.  
- **Product-Details**: Stores product data.  
- **logs**: Audit trail of all operations.  

### **Special Function**:  
- `JALALI()` converts Gregorian dates to Persian (Jalali) format.  

This script is ideal for small businesses tracking inventory in Persian.
