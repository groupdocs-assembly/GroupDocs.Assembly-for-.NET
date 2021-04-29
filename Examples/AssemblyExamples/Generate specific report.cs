using System;
using System.Collections.Generic;
using GroupDocs.Assembly;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class GenerateSpecificReport : AssemblyExamplesBase
    {
        [Test]
        public void LazilyAndRecursively()
        {
            //ExStart:LazilyAndRecursively
            DynamicEntity entity = new DynamicEntity(Guid.NewGuid());

            DocumentAssembler assembler = new DocumentAssembler();
            // Call AssembleDocument to generate Single Row Report in open document format.
            assembler.AssembleDocument(TemplatesDir + "Lazy and recursive.docx",
                ArtifactsDir + "GenerateSpecificReport.LazilyAndRecursively.docx", 
                new DataSourceInfo(entity, "root"));
            //ExEnd:LazilyAndRecursively
        }

        //ExStart:DynamicEntity
        public interface IPropertyProvider<out T>
        {
            T this[string propertyName] { get; }
        }

        public class DynamicEntity : IPropertyProvider<string>
        {
            /// <summary>
            /// Gets a property value by its name.
            /// </summary>
            public string this[string propertyName] =>
                // In this example, we simply return a property name as its value.
                // In a real-life application, a real property value should be returned.
                // This value can be cached using for example a Dictionary, or fetched
                // every time the property is requested.
                propertyName + " Value";

            /// <summary>
            /// Provides access to individual related <see cref="DynamicEntity"/> instances
            /// by their names.
            /// </summary>
            public IPropertyProvider<DynamicEntity> Entities => mEntities;

            /// <summary>
            /// Provides access to enumerations of related <see cref="DynamicEntity"/> instances
            /// by their names.
            /// </summary>
            public IPropertyProvider<IEnumerable<DynamicEntity>> Children => mChildren;

            private class ReferencedEntities : IPropertyProvider<DynamicEntity>
            {
                public DynamicEntity this[string propertyName] =>
                    // In this example, we simply return the root entity.
                    // In a real-life application, a DynamicEntity instance corresponding
                    // to propertyName for the given root entity should be returned.
                    // This instance can be cached using for example a Dictionary,
                    // or fetched every time the referenced entity is requested.
                    mRootEntity;

                public ReferencedEntities(DynamicEntity rootEntity)
                {
                    // The reference to the root entity allows to access fields of the root entity
                    // (such as an identifier) in the above indexer for a real-life application.
                    mRootEntity = rootEntity;
                }

                private readonly DynamicEntity mRootEntity;
            }

            private class ChildEntities : IPropertyProvider<IEnumerable<DynamicEntity>>
            {
                public IEnumerable<DynamicEntity> this[string propertyName]
                {
                    get
                    {
                        // In this example, we simply return the root entity three times.
                        // In a real-life application, an enumeration of DynamicEntity instances
                        // corresponding to propertyName for the given root entity should be returned.
                        // This enumeration can be cached using for example a Dictionary,
                        // or fetched every time the child entities are requested.
                        yield return mRootEntity;
                        yield return mRootEntity;
                        yield return mRootEntity;
                    }
                }

                public ChildEntities(DynamicEntity rootEntity)
                {
                    // The reference to the root entity allows to access fields of the root entity
                    // (such as an identifier) in the above indexer for a real-life application.
                    mRootEntity = rootEntity;
                }

                private readonly DynamicEntity mRootEntity;
            }
            public DynamicEntity(Guid id)
            {
                // In this example, we use Guid to represent an entity identifier.
                // In a real-life application, the identifier can be of any type or even missing.
                mId = id;

                // In this example, we simply initialize fields in the constructor.
                // In a real-life application, these fields can be initialized lazily
                // at the corresponding properties, if needed.
                mEntities = new ReferencedEntities(this);
                mChildren = new ChildEntities(this);
            }

            private readonly Guid mId;
            private readonly ReferencedEntities mEntities;
            private readonly ChildEntities mChildren;
        }
        //ExEnd:DynamicEntity
    }
}