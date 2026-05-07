# Vendor Unit Price Manager Security Specification

## Data Invariants
1. A vendor must have a name, email, and password.
2. Prices are sub-resources of a vendor and must link to a valid vendorId.
3. Vendors are sorted alphabetically in the UI.

## Access Control Logic
1. **Public Read**: The list of vendors (name, phone, fax, etc.) is public to allow users to select their company.
2. **Private Access**: Access to the unit price table is "gated" by a password check in the application.
3. **Admin Operations**: Adding/Editing vendors is currently open for the prototype, but should be restricted to authenticated admin users.

## The "Dirty Dozen" Payloads (Denial Tests)
1. Vendor without a password.
2. Price item with a negative price.
3. Price item with a 1MB string in `itemName`.
4. Updating a vendor's `name` to an empty string.
5. Deleting a vendor from a non-admin account (if admin is enabled).
6. Injecting a "ghost field" `isAdmin: true` into a user document.
7. Creating a price item for a non-existent vendor.
8. Rapidly creating 1000 vendors (rate limiting).
9. Modifying `createdAt` during an update.
10. Accessing prices of Vendor A using Vendor B's ID.
11. Providing an empty vendor name.
12. Using an ID with special characters like `../` to escape paths.
