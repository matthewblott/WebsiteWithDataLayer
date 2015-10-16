<%@ Control %>

<div>
  Enter "matt" for the username and "password" for the password and click "Login"
</div>

<form method="post" action="api/login">
  <input placeholder="Username" name="name" required />
  <input type="password" placeholder="Password" name="password" required />
  <input type="submit" value="Login" />
</form>