using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using EnvDTE;
using EnvDTE80;

namespace YongFa365.BatchFormat
{
    public static class ConfigHelper
    {

        public static List<string> ExcludePathList { get; } = new List<string> { "AssemblyInfo.cs", ".designer.cs", "Reference.cs", @"\MetaData\" };

        public static List<string> GetExcludePathList(this DTE2 input)
        {
            var temp = input.Properties["BatchFormat", "General"].Item("ExcludePath").Value;
            if (string.IsNullOrWhiteSpace(temp?.ToString()))
            {
                return ExcludePathList;
            }
            var result = temp.ToString().Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries).ToList();
            return result;
        }

        public static bool IsIgnoreT4Child(this DTE2 input)
        {
            var temp = input.Properties["BatchFormat", "General"].Item("IgnoreT4Child").Value;
            if (string.IsNullOrWhiteSpace(temp?.ToString()))
            {
                return true;
            }
            return (bool)temp;
        }
    }

    public static class EnumerableExtensions
    {
        public static void ForEach<TItem>(this IEnumerable<TItem> source, Action<TItem> action)
        {
            foreach (var local in source)
            {
                action(local);
            }
        }
    }

    public sealed class ProjectIterator : IEnumerable<Project>
    {
        private readonly Solution _solution;

        public ProjectIterator(Solution solution)
        {
            if (solution == null)
            {
                throw new ArgumentNullException(nameof(solution));
            }
            _solution = solution;
        }

        private IEnumerable<Project> Enumerate(Project project)
        {
            var enumerator = project.ProjectItems.GetEnumerator();
            while (enumerator.MoveNext())
            {
                var current = (ProjectItem)enumerator.Current;
                var currentObject = current.Object as Project;

                if (currentObject != null)
                {
                    if (currentObject.Kind != "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}")
                    {
                        yield return currentObject;
                    }
                    else
                    {
                        foreach (var iteratorVariable1 in Enumerate(currentObject))
                        {
                            yield return iteratorVariable1;
                        }
                    }
                }
            }
        }

        public IEnumerator<Project> GetEnumerator()
        {
            var enumerator = _solution.Projects.GetEnumerator();
            while (enumerator.MoveNext())
            {
                var current = (Project)enumerator.Current;
                if (current.Kind != "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}")
                {
                    yield return current;
                }
                else
                {
                    foreach (var iteratorVariable1 in Enumerate(current))
                    {
                        yield return iteratorVariable1;
                    }
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }


    }

    public sealed class ProjectItemIterator : IEnumerable<ProjectItem>
    {
        private readonly ProjectItems _items;

        public ProjectItemIterator(ProjectItems items)
        {
            if (items == null)
            {
                throw new ArgumentNullException(nameof(items));
            }
            _items = items;
        }

        private IEnumerable<ProjectItem> Enumerate(ProjectItems items)
        {
            var enumerator = items.GetEnumerator();
            while (enumerator.MoveNext())
            {
                var current = (ProjectItem)enumerator.Current;
                yield return current;
                foreach (var iteratorVariable1 in Enumerate(current.ProjectItems))
                {
                    yield return iteratorVariable1;
                }
            }
        }

        public IEnumerator<ProjectItem> GetEnumerator()
        {
            return Enumerate(_items)
                  .Select(item => item)
                  .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }


    }
}
