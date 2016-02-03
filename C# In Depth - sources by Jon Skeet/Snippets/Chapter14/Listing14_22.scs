dynamic size = 5;
var numbers = Enumerable.Range(10, 10);
// Error: extension methods can't be dynamically dispatched
// var error = numbers.Take(size);
var workaround1 = numbers.Take((int) size);
var workaround2 = Enumerable.Take(numbers, size);