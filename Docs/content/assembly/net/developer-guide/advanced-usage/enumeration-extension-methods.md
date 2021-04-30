---
id: enumeration-extension-methods
url: assembly/net/enumeration-extension-methods
title: Enumeration Extension Methods
weight: 2
description: ""
keywords: 
productName: GroupDocs.Assembly for .NET
hideChildren: False
---
LINQ Reporting Engine enables you to perform common manipulations on a sequential data through the engine's built-in extension methods for `IEnumerable`. These extension methods mimic some extension methods of [IEnumerable&lt;T&gt;](http://msdn.microsoft.com/en-us/library/9eekhta0\(v=vs.110\).aspx) providing the same signatures and behavior features. Thus you can group, sort and perform other sequential data manipulations in template expressions in a familiar way.

The following table describes these built-in extension methods. The following notation conventions are used within the table:

- Selector stands for a lambda function returning a value and taking an enumeration item as its single argument. See “Using Lambda Functions” for more information.
- `ComparableSelector` stands for `Selector` returning [IComparable](http://msdn.microsoft.com/en-US/library/system.icomparable\(v=vs.110\).aspx).
- `Predicate` stands for `Selector` returning a Boolean value.

Examples in the following table are given using persons and `otherPersons`, enumerations of instances of the `Person` class that is defined as follows.

{{< highlight csharp >}}
public class Person
{
    public String Name { get { ... } }
    public int Age { get { ... } }
    public IEnumerable<Person> Children { get { ... } }
    ...
}
{{< /highlight >}}

|Extension Method |Examples and Notes |
| :- | :- |
|All(Predicate)|{{< highlight csharp >}}
persons.All(p => p.Age < 50){{< /highlight >}}|
|Any()|{{< highlight csharp >}}
persons.Any(){{< /highlight >}}|
|Any(Predicate)|{{< highlight csharp >}}
persons.Any(p => p.Name == "John Smith"){{< /highlight >}}|
|Average(Selector)|{{< highlight csharp >}}
persons.Average(p => p.Age){{< /highlight >}}The input selector must return a value of any type that has predefined or user-defined addition and division operators. |
|Concat(IEnumerable)|{{< highlight csharp >}}
persons.Concat(otherPersons){{< /highlight >}}An implicit reference conversion must exist between types of items of concatenated enumerations. |
|Contains(Object)|{{< highlight csharp >}}
persons.Contains(otherPersons.First()){{< /highlight >}}|
|Count()|{{< highlight csharp >}}
persons.Count(){{< /highlight >}}|
|Count(Predicate)|{{< highlight csharp >}}
persons.Count(p => p.Age > 30){{< /highlight >}}|
|Distinct()|{{< highlight csharp >}}
persons.Distinct(){{< /highlight >}}|
|First()|{{< highlight csharp >}}
persons.First(){{< /highlight >}}|
|First(Predicate)|{{< highlight csharp >}}
persons.First(p => p.Age > 30){{< /highlight >}}|
|FirstOrDefault()|{{< highlight csharp >}}
persons.FirstOrDefault(){{< /highlight >}}|
|FirstOrDefault(Predicate)|{{< highlight csharp >}}
persons.FirstOrDefault(p => p.Age > 30){{< /highlight >}}|
|GroupBy(Selector)|{{< highlight csharp >}}
persons.GroupBy(p => p.Age){{< /highlight >}}Or {{< highlight csharp >}}
persons.GroupBy(    p => new { Age = p.Age, Count = p.Children.Count() }){{< /highlight >}}This method returns an enumeration of group objects. Each group has a unique key defined by the input selector and contains items of the source enumeration associated with this key. You can access the key of a group instance using the Key property. You can treat a group itself as an enumeration of items that the group contains. |
|Last()|{{< highlight csharp >}}
persons.Last(){{< /highlight >}}|
|Last(Predicate)|{{< highlight csharp >}}
persons.Last(p => p.Age > 100){{< /highlight >}}|
|LastOrDefault()|{{< highlight csharp >}}
persons.LastOrDefault(){{< /highlight >}}|
|LastOrDefault(Predicate)|{{< highlight csharp >}}
persons.LastOrDefault(p => p.Age > 100){{< /highlight >}}|
|Max(ComparableSelector)|{{< highlight csharp >}}
persons.Max(p => p.Age){{< /highlight >}}|
|Min(ComparableSelector)|{{< highlight csharp >}}
persons.Min(p => p.Age){{< /highlight >}}|
|OrderBy(ComparableSelector)|{{< highlight csharp >}}
persons.OrderBy(p => p.Age){{< /highlight >}}Or {{< highlight csharp >}}
persons.OrderBy(p => p.Age).ThenByDescending(p => p.Name){{< /highlight >}}Or {{< highlight csharp >}}
persons.OrderBy(p => p.Age).ThenByDescending(p => p.Name).ThenBy(p => p.Children.Count()){{< /highlight >}}This method returns an enumeration ordered by a single key. To specify additional ordering keys, you can use the following extension methods of an ordered enumeration:- ThenBy(ComparableSelector)- ThenByDescending(ComparableSelector)|
|OrderByDescending(ComparableSelector)|{{< highlight csharp >}}
persons.OrderByDescending(p => p.Age){{< /highlight >}}Or {{< highlight csharp >}}
persons.OrderByDescending(p => p.Age).ThenByDescending(p => p.Name){{< /highlight >}}Or {{< highlight csharp >}}
persons.OrderByDescending(p => p.Age).ThenByDescending(p => p.Name).ThenBy(p => p.Children.Count()){{< /highlight >}}See the previous note. |
|Select(Selector)|{{< highlight csharp >}}
persons.Select(p => p.Name){{< /highlight >}}|
|SelectMany(EnumerationSelector)|{{< highlight csharp >}}
persons.SelectMany(p => p.Children){{< /highlight >}}|
|Single()|{{< highlight csharp >}}
persons.Single(){{< /highlight >}}|
|Single(Predicate)|{{< highlight csharp >}}
persons.Single(p => p.Name == "John Smith"){{< /highlight >}}|
|SingleOrDefault()|{{< highlight csharp >}}
persons.SingleOrDefault(){{< /highlight >}}|
|SingleOrDefault(Predicate)|{{< highlight csharp >}}
persons.SingleOrDefault(p => p.Name == "John Smith"){{< /highlight >}}|
|Skip(int)|{{< highlight csharp >}}
persons.Skip(10){{< /highlight >}}|
|SkipWhile(Predicate)|{{< highlight csharp >}}
persons.SkipWhile(p => p.Age < 21){{< /highlight >}}|
|Sum(Selector)|{{< highlight csharp >}}
persons.Sum(p => p.Children.Count()){{< /highlight >}}The input selector must return a value of any type that has a predefined or user-defined addition operator. |
|Take(int)|{{< highlight csharp >}}
persons.Take(5){{< /highlight >}}|
|TakeWhile(Predicate)|{{< highlight csharp >}}
persons.TakeWhile(p => p.Age < 50){{< /highlight >}}|
|Union(IEnumerable)|{{< highlight csharp >}}
persons.Union(otherPersons){{< /highlight >}}An implicit reference conversion must exist between types of items of united enumerations. |
|Where(Predicate)|{{< highlight csharp >}}
persons.Where(p => p.Age > 18){{< /highlight >}}|