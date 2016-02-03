string text = "cut me up";
dynamic guid = Guid.NewGuid();
text.Substring(guid);
// Compile-time error: no overload with string and something else
// text.Substring("x", guid);
// Compile-time error: no overload with three parameters
// text.Substring(guid, guid, guid);