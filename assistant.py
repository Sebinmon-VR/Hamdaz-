
def assistant_decorator(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for('login'))
        user = session.get("user")
        return f(user, *args, **kwargs)
    return decorated_function

