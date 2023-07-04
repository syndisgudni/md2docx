## X Short Description

| | |
| ----------------- | ------- |
| **Severity**          | **Critical / High / Medium / Low / Informational** |
| **Date Sent**         | 2023-07-04 |
| **Location**          | `src/code/index.php#1337` |
| **Issue Identifier**  | `SYN-XX-Y-Z` |
| **Impact**            | Comments sent by X to Y are not sanitized, leading to a cross-site scripting vulnerability when a user views an order in Y. |

### Technical Details

\<Description\>

```javascript
function fibonacci(num) {
  var a = 1,
    b = 0,
    temp;

  while (num >= 0) {
    temp = a;
    a = a + b;
    b = temp;
    num++;
  }

  return b;
}
```
**Figure X:** *Code figure.*

```bash
user@hostname:/home/user/docs/secrets$ ls -lah
drwxr-xr-x user user 4.0 KB Fri Apr 16 15:35:15 2021 .
drwx------ user user  20 KB Wed Apr 28 16:10:35 2021 ..
drwxr-xr-x user user 4.0 KB Mon Mar 15 15:34:48 2021 more_secrets
drwxr-xr-x user user 4.0 KB Mon Jan 11 22:13:33 2021 even_more_secrets
.rw-r--r-- user user 1.1 KB Mon Jan 25 09:51:56 2021 secrets.json
[...]
```
**Figure X:** *Shell figure.*

```reqres
--- Request
POST /api/foo HTTP/1.1
Host: syndis.is
[...]

{"kennitala":"1234561111","dateOfBirth":"1956-12-34T00:00:00.000Z","nafn":"syndis","netfang":"syndis@syndis.is"}
--- Response
HTTP/1.1 200 OK
[...]

Very Serious Information
```
**Figure X:** *Reproduction request/response.*

![puppy](puppy.jpeg)
**Figure X:** *Image figure.*

| **Column 1** | **Column 2** | **Column 3** |
| ------------ | ------------ | ------------ |
| One          | Two          | Three        |
| Four         | Five         | Six          |
| Seven        | Eight        | Nine         |
**Figure X:** *Fancy table.*

### Recommendation

\<Recommendation\>
