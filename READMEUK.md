Are you tired of creating your Active Directory users one by one, risking mistakes with the wrong Organizational Unit (OU), or forgetting to assign the correct groups and Office 365 licenses? I have just the PowerShell script you need!

Below is an overview of its main features:

    Administrator Rights Check
    To protect your environment, the script immediately verifies your permissions. If you lack the required privileges, account creation is blocked right away.

    Active Directory User Creation
    Say goodbye to repetitive manual tasks in the GUI: simply enter the user’s first name, last name, department, phone number, and password. The script then automatically:
        Generates a unique SamAccountName, preventing any conflicts.
        Creates the UPN and DisplayName.
        Fills in essential fields such as department, city, address, manager, etc.
        Places the user in the correct OU, determined dynamically.

    Phone Numbers and Contact Information Management
    The script accounts for both mobile and landline phone numbers, validating formats (for instance, 06 12 34 56 78).

    Group Assignment
    Depending on the workstation type (desktop, laptop, or thin client), the script automatically adds the user to the relevant groups:
        Desktop with or without a mobile device.
        Laptop, including VPN groups if needed.
        Thin client.

    Inheriting Permissions from Another User
    If you want a new hire to mirror the same configuration as an existing colleague (same group memberships, same privileges, except for critical roles), simply select the reference user, and the script transfers those permissions seamlessly.

    Managing Email Aliases
    Easily add extra SMTP aliases by choosing from available domains. This step is optional and entirely up to your needs.

    Office 365 License Assignment
    The script presents labeled groups (E3, Basic, etc.): pick the license you need, and the user is automatically placed in the corresponding group. Additional products, such as Visio Plan2, can also be added with a single click.

    Optional Menu and Modular Code
        A main menu (Show-Menu) brings all functionalities together in one place.
        Dedicated functions (Create-User, Test-AdminAccess, etc.) let you tailor the script to your environment.
        Submenus handle tasks like managing aliases, Office 365 license packs, or viewing modification history.
        A CSV import feature fully automates user creation (no more manual copy-pasting).

    Securing Login Attempts
    To strengthen security, the script limits administrator authentication attempts to three, effectively preventing unauthorized access.

    Future-Proof and Customizable
    Want to add a new field or require a specific phone format? The script is thoroughly commented so you can adapt it to your exact needs.

In Summary

    Eliminate repetitive tasks in the graphical interface.
    Automate and secure user creation, group assignments, attribute updates, and Office 365 licensing.
    Reduce mistakes and oversights with a clear, guided process.

This PowerShell script handles nearly everything for you, with selection dialogs to minimize typos. In short, it saves you valuable time and spares you many administrative headaches. Ready to streamline those repetitive tasks? Run the script and enjoy its many benefits!
