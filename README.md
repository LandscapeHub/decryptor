# Decryptor

This is a tool used to decrypt password protected xlsx files using the command line

## Building

Simply run `mvn package`.
You will see the fully functional and portable jar: `./target/decryptor.jar`

## Usage
1. Move the jar to the same directory as your encrypted xlsx file
2. Run the java jar with 3 arguments, input file, output file, and password.

    Example:

    `java -jar decryptor.jar encryptedFile.xlsx unlocked.xlsx My$ecretpassw0rd101`

3. You now should be able to open the unlocked xlsx file without a password
