using System;
using System.Linq;
using System.Reflection;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.MasterData;
using Eplan.EplApi.Scripting;



namespace LabelModificationActions
{
    public class ResolveExternalDocumentPath
    {
        /// <summary>
        /// Call from labeling action (single label)
        /// </summary>
        /// <param name="context">Has to have DocumentIndex as parameter</param>
        [DeclareAction("ResolveExternalDocumentPathAction")]
        public void Execute(ActionCallingContext context)
        {
            // Get the database Ids from labeling interface
            var labelObjects = string.Empty;
            context.GetParameter("objects", ref labelObjects);
            var databaseId = labelObjects.Split(';')[0]; // Summarized parts lists contains MergedArticleReferences
            
            // Get the MergedArticleReference from string identifier by reflection
            var mergedArticleReference = StorableObjectWrapper.FromStringIdentifier(databaseId);
            
            // Wrap the MergedArticleReference it into a custom runtime type to get the ArticleReference from it
            var mergedArticleReferenceWrapper = new MergedArticleReferenceWrapper(mergedArticleReference);
            
            // Get the ArticleReference from MergedArticleReference by reflection
            var articleReference = mergedArticleReferenceWrapper.GetMainArticleReference();
            
            // Wrap the ArticleReference it into a custom runtime type to get the part infos from it
            var articleReferenceWrapper = new ArticleReferenceWrapper(articleReference);
            
            // Get part infos
            var partNr = articleReferenceWrapper.PartNr;
            var variantNr = articleReferenceWrapper.VariantNr;

            // Get the document index that should be returned 
            var documentIndexString = string.Empty;
            context.GetParameter("DocumentIndex", ref documentIndexString);
            var documentIndex = int.Parse(documentIndexString);

            // Get external document link from parts database
            var articleExternalDocumentValue = string.Empty;
            using (var partsDb = new MDPartsManagement().OpenDatabase())
            {
                var part = partsDb.GetPart(partNr, variantNr);
                if (part == null)
                {
                    // Set a message to the label
                    context.SetStrings(new []{string.Format("Part not found in database Part No: {0} Variant: {1}", partNr, variantNr)});
                    return;
                }
                var externalDocumentValue = part.Properties.ARTICLE_EXTERNAL_DOCUMENT[documentIndex];
                if (!externalDocumentValue.IsEmpty)
                {
                    articleExternalDocumentValue = externalDocumentValue;
                }
            }

            // The indexed property comes with the description separated by line break
            articleExternalDocumentValue = articleExternalDocumentValue.Split('\n')[0];
            
            // Remove the eplan specific path variable 
            articleExternalDocumentValue = PathMap.SubstitutePath(articleExternalDocumentValue);
            
            // Set the final value to the label 
            context.SetStrings(new []{articleExternalDocumentValue});
        }
    }

    public class ArticleReferenceWrapper
    {
        private readonly object _articleReference;
        
        public ArticleReferenceWrapper(object articleReference)
        {
            _articleReference = articleReference;
        }

        public string PartNr
        {
            get
            {
                var articleReferenceType = _articleReference.GetType();
                var partNrProperty = articleReferenceType.GetProperty("PartNr", BindingFlags.Instance | BindingFlags.Public);
                var partNr =  partNrProperty.GetValue(_articleReference);
                return partNr.ToString();
            }
        }
        
        public string VariantNr
        {
            get
            {
                var articleReferenceType = _articleReference.GetType();
                var variantNrProperty = articleReferenceType.GetProperty("VariantNr", BindingFlags.Instance | BindingFlags.Public);
                var variantNr =  variantNrProperty.GetValue(_articleReference);
                return variantNr.ToString();
            }
        }
    }
    
    public class MergedArticleReferenceWrapper
    {
        private readonly object _mergedArticleReference;

        public MergedArticleReferenceWrapper(object mergedArticleReference)
        {
            _mergedArticleReference = mergedArticleReference;
        }
        
        public object GetMainArticleReference()
        {
            var mergedArticleReferenceType = _mergedArticleReference.GetType();
            var getMainArticleReference = mergedArticleReferenceType.GetMethod("GetMainArticleReference");
            var articleReference = getMainArticleReference.Invoke(_mergedArticleReference, new object[] { });
            return articleReference;
        }
    }

    public class StorableObjectWrapper
    {
        private static Assembly _dataModelAssembly;
        
        public static object FromStringIdentifier(string databaseId)
        {
            if (_dataModelAssembly == null)
            {
                _dataModelAssembly = AppDomain.CurrentDomain.GetAssemblies()
                    .FirstOrDefault(a => a.FullName.StartsWith("Eplan.EplApi.DataModelu"));
            }
            
            var storableObjectType = _dataModelAssembly.ExportedTypes.FirstOrDefault(t => t.Name == "StorableObject");
            MethodInfo fromStringIdentifier = storableObjectType.GetMethod("FromStringIdentifier",
                BindingFlags.Public | BindingFlags.Static, null, new[] { typeof(string) }, null);
            var args = new object[]{databaseId};
            var storableObject = fromStringIdentifier.Invoke(null, args);
            return storableObject;
        }
    }
}