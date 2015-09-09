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

        public static readonly List<string> ExcludePathList = new List<string> { "AssemblyInfo.cs", ".designer.cs", "Reference.cs", "/metadata/" };

        public static List<string> GetValue(this DTE2 input, string key)
        {
            var temp = input.Properties["BatchFormat", "General"].Item(key).Value;
            if (string.IsNullOrWhiteSpace(temp?.ToString()))
            {
                return ExcludePathList;
            }
            return temp.ToString().Split(new[] { "|" }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }
    }

    public static class EnumerableExtensions
    {
        public static void ForEach<TItem>(this IEnumerable<TItem> source, Action<TItem> action)
        {
            foreach (TItem local in source)
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
                throw new ArgumentNullException("solution");
            }
            _solution = solution;
        }

        private IEnumerable<Project> Enumerate(Project project)
        {
            IEnumerator enumerator = project.ProjectItems.GetEnumerator();
            while (enumerator.MoveNext())
            {
                ProjectItem current = (ProjectItem)enumerator.Current;
                Project currentObject = current.Object as Project;

                if (currentObject != null)
                {
                    if (currentObject.Kind != "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}")
                    {
                        yield return currentObject;
                    }
                    else
                    {
                        foreach (Project iteratorVariable1 in Enumerate(currentObject))
                        {
                            yield return iteratorVariable1;
                        }
                    }
                }
            }
        }

        public IEnumerator<Project> GetEnumerator()
        {
            IEnumerator enumerator = _solution.Projects.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Project current = (Project)enumerator.Current;
                if (current.Kind != "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}")
                {
                    yield return current;
                }
                else
                {
                    foreach (Project iteratorVariable1 in Enumerate(current))
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
                throw new ArgumentNullException("items");
            }
            _items = items;
        }

        private IEnumerable<ProjectItem> Enumerate(ProjectItems items)
        {
            IEnumerator enumerator = items.GetEnumerator();
            while (enumerator.MoveNext())
            {
                ProjectItem current = (ProjectItem)enumerator.Current;
                yield return current;
                foreach (ProjectItem iteratorVariable1 in Enumerate(current.ProjectItems))
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
