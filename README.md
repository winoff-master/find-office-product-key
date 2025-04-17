# How to Retrieve Your Office Product Key
Lost your Office product key? You can try this method:
1. Open **PowerShell** as administrator.
2. Enter the following command:
```
(Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey
```
3. If your key is stored in BIOS, it will appear.
> **Note:** This method may not work with digital licenses from Microsoft Store or OEM keys.
---
find-office-product-key