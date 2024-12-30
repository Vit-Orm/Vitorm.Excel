using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Tasks;

using OfficeOpenXml;

using Vit.Linq;
using Vit.Linq.FilterRules;
using Vit.Linq.FilterRules.ComponentModel;

using Vitorm.Entity;

namespace Vitorm.Excel
{
    // https://epplussoftware.com/en/Developers/
    // https://github.com/EPPlusSoftware/EPPlus.Samples.CSharp

    public class DbSetConstructor
    {
        public static IDbSet CreateDbSet(IDbContext dbContext, IEntityDescriptor entityDescriptor)
        {
            return _CreateDbSet.MakeGenericMethod(entityDescriptor.entityType, entityDescriptor.key?.type ?? typeof(string))
                     .Invoke(null, new object[] { dbContext, entityDescriptor }) as IDbSet;
        }

        static readonly MethodInfo _CreateDbSet = new Func<DbContext, IEntityDescriptor, IDbSet>(CreateDbSet<object, string>)
                   .Method.GetGenericMethodDefinition();
        public static IDbSet<Entity> CreateDbSet<Entity, EntityKey>(DbContext dbContext, IEntityDescriptor entityDescriptor)
        {
            return new DbSet<Entity, EntityKey>(dbContext, entityDescriptor);
        }

    }


    public partial class DbSet<Entity, EntityKey> : IDbSet<Entity>
    {
        public virtual IDbContext dbContext { get; protected set; }
        public virtual DbContext DbContext => (DbContext)dbContext;


        protected IEntityDescriptor _entityDescriptor;
        public virtual IEntityDescriptor entityDescriptor => _entityDescriptor;


        public DbSet(DbContext dbContext, IEntityDescriptor entityDescriptor)
        {
            this.dbContext = dbContext;
            this._entityDescriptor = entityDescriptor;
        }

        // #0 Schema :  ChangeTable
        public virtual IEntityDescriptor ChangeTable(string tableName) => _entityDescriptor = _entityDescriptor.WithTable(tableName);
        public virtual IEntityDescriptor ChangeTableBack() => _entityDescriptor = _entityDescriptor.GetOriginEntityDescriptor();


        public virtual ExcelPackage package => DbContext.dbConnection;
        public virtual ExcelWorksheet sheet => package.Workbook.Worksheets[entityDescriptor.tableName];



        #region columnIndexs
        Dictionary<string, int> _columnIndexes;

        Dictionary<string, int> columnIndexes =>
            _columnIndexes ??=
                Enumerable.Range(1, sheet.Dimension?.End.Column ?? 0)
                .Select(i => (index: i, columnName: sheet.GetValue<string>(1, i)))
                .GroupBy(m => m.columnName)
                .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                .Select(g => g.First())
                .ToDictionary(item => item.columnName, item => item.index)
            ;
        #endregion



        #region Excel Methods
        protected virtual void Save()
        {
            _columnIndexes = null;
            package.Save();
        }
        protected virtual async Task SaveAsync()
        {
            _columnIndexes = null;
            await package.SaveAsync();
        }



        protected virtual int GetMaxId()
        {
            var entities = GetEntities().Select(m => m.entity);
            if (entities?.Any() != true) return 0;
            return entities.Max(entity => int.TryParse(entityDescriptor.key.GetValue(entity)?.ToString(), out var id) ? id : 0);
        }


        public virtual void AddColumnsIfNotExist()
        {
            var colsToAdd = entityDescriptor.properties.Where(col => !columnIndexes.TryGetValue(col.columnName, out var _)).ToList();

            if (!colsToAdd.Any()) return;

            int colIndex = sheet.Columns.EndColumn;
            foreach (var col in colsToAdd)
            {
                colIndex++;

                var column = sheet.Column(colIndex);
                sheet.SetValue(1, colIndex, col.columnName);
            }
            _columnIndexes = null;
        }


        public virtual void SetDateTimeFormat(string format = "yyyy-MM-dd HH:mm:ss")
        {
            foreach (var col in entityDescriptor.properties.Where(col => TypeUtil.GetUnderlyingType(col.type) == typeof(DateTime)))
            {
                if (!columnIndexes.TryGetValue(col.columnName, out var colIndex)) continue;

                var column = sheet.Columns[colIndex];

                column.Style.Numberformat.Format = format;
                column.AutoFit();
            }
        }


        protected virtual void SetRow(Entity entity, int rowIndex)
        {
            foreach (var col in entityDescriptor.properties)
            {
                if (!columnIndexes.TryGetValue(col.columnName, out var colIndex)) continue;
                var value = col.GetValue(entity);
                sheet.SetValue(rowIndex, colIndex, value);
            }
        }

        protected virtual int DeleteByKeysWithoutSave<Key>(IEnumerable<Key> keys)
        {
            IEnumerable<EntityKey> entityKeys;
            if (typeof(Key) == typeof(EntityKey))
            {
                entityKeys = (IEnumerable<EntityKey>)keys;
            }
            else
            {
                entityKeys = keys.Select(key => (EntityKey)TypeUtil.ConvertToType(key, typeof(EntityKey)));
            }

            int colIndex = columnIndexes.TryGetValue(entityDescriptor.key.columnName, out var i) ? i : throw new ArgumentOutOfRangeException("key column not exist.");

            var lastRowIndex = sheet.Dimension.End.Row;

            int count = 0;
            for (var rowIndex = lastRowIndex; rowIndex >= 2; rowIndex--)
            {
                var key = (EntityKey)TypeUtil.ConvertToType(sheet.GetValue(rowIndex, colIndex), typeof(EntityKey));
                if (entityKeys.Contains(key))
                {
                    sheet.DeleteRow(rowIndex);
                    count++;
                }
            }
            return count;
        }

        protected virtual IEnumerable<(int rowIndex, Entity entity)> GetEntities()
        {
            if (sheet?.Dimension == null) yield break;

            var lastRowIndex = sheet.Dimension.End.Row;
            for (var rowIndex = 2; rowIndex <= lastRowIndex; rowIndex++)
            {
                var entity = (Entity)Activator.CreateInstance(entityDescriptor.entityType);
                try
                {
                    foreach (var col in entityDescriptor.properties)
                    {
                        if (!columnIndexes.TryGetValue(col.columnName, out var colIndex)) continue;

                        if (col.type == typeof(DateTime) || col.type == typeof(DateTime?))
                        {
                            var value = sheet.GetValue<DateTime>(rowIndex, colIndex);
                            col.SetValue(entity, value);
                        }
                        else
                        {
                            var value = sheet.GetValue(rowIndex, colIndex);
                            value = TypeUtil.ConvertToType(value, col.type);
                            col.SetValue(entity, value);
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw;
                }
                yield return (rowIndex, entity);
            }

        }
        public virtual Expression<Func<Entity, bool>> GetKeyPredicate(object keyValue)
        {
            var filter = new FilterRule { field = entityDescriptor.key.propertyName, @operator = "=", value = keyValue };
            var predicate = FilterService.Instance.ConvertToCode_PredicateExpression<Entity>(filter);
            return predicate;
        }

        public virtual Expression<Func<Entity, bool>> GetKeyPredicate<Key>(IEnumerable<Key> keys)
        {
            var filter = new FilterRule { field = entityDescriptor.key.propertyName, @operator = "In", value = keys };
            var predicate = FilterService.Instance.ConvertToCode_PredicateExpression<Entity>(filter);
            return predicate;
        }
        #endregion




        #region #0 Schema :  Create Drop

        public virtual bool TableExist()
        {
            return sheet != null;
        }

        public virtual void TryCreateTable()
        {
            if (TableExist()) return;

            var sheet = package.Workbook.Worksheets.Add(entityDescriptor.tableName);

            int colIndex = 0;
            foreach (var col in entityDescriptor.properties)
            {
                colIndex++;

                var column = sheet.Column(colIndex);
                sheet.SetValue(1, colIndex, col.columnName);
            }

            SetDateTimeFormat();

            Save();
        }


        public virtual async Task TryCreateTableAsync()
        {
            if (TableExist()) return;

            var sheet = package.Workbook.Worksheets.Add(entityDescriptor.tableName);

            int colIndex = 0;
            foreach (var col in entityDescriptor.properties)
            {
                colIndex++;

                var column = sheet.Column(colIndex);
                sheet.SetValue(1, colIndex, col.columnName);
            }


            SetDateTimeFormat();

            await SaveAsync();
        }

        public virtual void TryDropTable()
        {
            if (!TableExist()) return;
            package.Workbook.Worksheets.Delete(entityDescriptor.tableName);
            Save();
        }
        public virtual async Task TryDropTableAsync()
        {
            if (!TableExist()) return;

            package.Workbook.Worksheets.Delete(entityDescriptor.tableName);
            await SaveAsync();
        }


        public virtual void Truncate()
        {
            var lastRowIndex = sheet.Dimension?.End.Row ?? 0;
            if (lastRowIndex < 2) return;

            sheet.DeleteRow(2, lastRowIndex);
            Save();
        }

        public virtual async Task TruncateAsync()
        {
            var lastRowIndex = sheet.Dimension?.End.Row ?? 0;
            if (lastRowIndex < 2) return;

            sheet.DeleteRow(2, lastRowIndex);
            await SaveAsync();
        }

        #endregion


        #region #1 Create :  Add AddRange
        public virtual Entity Add(Entity entity)
        {
            AddRange(new[] { entity });
            return entity;
        }


        public virtual async Task<Entity> AddAsync(Entity entity)
        {
            await AddRangeAsync(new[] { entity });
            return entity;
        }
        public virtual void AddRange(IEnumerable<Entity> entities)
        {
            AddColumnsIfNotExist();


            #region generate identity key if needed
            if (entityDescriptor.key.isIdentity)
            {
                int maxId = GetMaxId();

                entities.ForEach(entity =>
                {
                    object keyValue = entityDescriptor.key.GetValue(entity);
                    var keyIsEmpty = keyValue is null || keyValue.Equals(TypeUtil.GetDefaultValue(entityDescriptor.key.type));
                    if (keyIsEmpty)
                    {
                        maxId++;
                        entityDescriptor.key.SetValue(entity, maxId);
                    }
                });
            }
            #endregion


            var lastRowIndex = sheet.Dimension?.End.Row ?? 0;
            var range = sheet.Cells[lastRowIndex + 1, 1];

            //range.LoadFromCollection(entities, PrintHeaders: false);
            //var dictionaries = entities.Select(entity => entityDescriptor.allColumns.ToDictionary(col => col.columnName, col => col.GetValue(entity)));
            //range.LoadFromDictionaries(dictionaries, printHeaders: false);

            foreach (var entity in entities)
            {
                lastRowIndex++;
                SetRow(entity, lastRowIndex);
            }

            Save();
        }


        public virtual async Task AddRangeAsync(IEnumerable<Entity> entities)
        {
            AddColumnsIfNotExist();


            #region generate identity key if needed
            if (entityDescriptor.key.isIdentity)
            {
                int maxId = GetMaxId();

                entities.ForEach(entity =>
                {
                    object keyValue = entityDescriptor.key.GetValue(entity);
                    var keyIsEmpty = keyValue is null || keyValue.Equals(TypeUtil.GetDefaultValue(entityDescriptor.key.type));
                    if (keyIsEmpty)
                    {
                        maxId++;
                        entityDescriptor.key.SetValue(entity, maxId);
                    }
                });
            }
            #endregion


            var lastRowIndex = sheet.Dimension?.End.Row ?? 0;
            var range = sheet.Cells[lastRowIndex + 1, 1];

            //range.LoadFromCollection(entities, PrintHeaders: false);
            //var dictionaries = entities.Select(entity => entityDescriptor.allColumns.ToDictionary(col => col.columnName, col => col.GetValue(entity)));
            //range.LoadFromDictionaries(dictionaries, printHeaders: false);

            foreach (var entity in entities)
            {
                lastRowIndex++;
                SetRow(entity, lastRowIndex);
            }

            await SaveAsync();
        }





        #endregion


        #region #2 Retrieve : Get Query

        public virtual Entity Get(object keyValue)
        {
            var predicate = GetKeyPredicate(keyValue);
            return Query().FirstOrDefault(predicate);
        }

        public virtual Task<Entity> GetAsync(object keyValue)
        {
            return Task.Run(() => Get(keyValue));
        }

        public virtual IQueryable<Entity> Query()
        {
            return GetEntities().Select(m => m.entity).AsQueryable();
        }

        #endregion


        #region #3 Update: Update UpdateRange
        public virtual int Update(Entity entity)
        {
            int count = UpdateWithoutSave(entity);
            Save();
            return count;
        }
        public virtual async Task<int> UpdateAsync(Entity entity)
        {
            int count = UpdateWithoutSave(entity);
            await SaveAsync();
            return count;
        }

        protected virtual int UpdateWithoutSave(Entity entity)
        {
            AddColumnsIfNotExist();

            var key = entityDescriptor.key.GetValue(entity);
            int count = 0;
            foreach (var item in GetEntities())
            {
                var oldEntity = item.entity;
                var oldKey = entityDescriptor.key.GetValue(oldEntity);
                if (!key.Equals(oldKey)) continue;

                var rowIndex = item.rowIndex;
                SetRow(entity, rowIndex);
                count++;
                break;
            }
            return count;
        }


        public virtual int UpdateRange(IEnumerable<Entity> entities)
        {
            int count = UpdateRangeWithoutSave(entities);
            Save();
            return count;
        }
        public virtual async Task<int> UpdateRangeAsync(IEnumerable<Entity> entities)
        {
            int count = UpdateRangeWithoutSave(entities);
            await SaveAsync();
            return count;
        }


        protected virtual int UpdateRangeWithoutSave(IEnumerable<Entity> entities)
        {
            AddColumnsIfNotExist();

            // key -> entity
            var entityMap =
                 entities.Select(entity => (key: (EntityKey)entityDescriptor.key.GetValue(entity), entity: entity))
                 .GroupBy(item => item.key).Select(group => (key: group.Key, entity: group.Last().entity))
                 .ToDictionary(item => item.key, item => item.entity);

            int count = 0;
            foreach (var item in GetEntities())
            {
                var oldEntity = item.entity;
                var key = (EntityKey)entityDescriptor.key.GetValue(oldEntity);
                if (!entityMap.TryGetValue(key, out var entity)) continue;

                var rowIndex = item.rowIndex;
                SetRow(entity, rowIndex);
                count++;
            }
            return count;
        }

        #endregion


        #region #4 Delete : Delete DeleteRange DeleteByKey DeleteByKeys

        public virtual int Delete(Entity entity)
        {
            var keyValue = entityDescriptor.key.GetValue(entity);
            return DeleteByKey(keyValue);
        }
        public virtual Task<int> DeleteAsync(Entity entity)
        {
            var keyValue = entityDescriptor.key.GetValue(entity);
            return DeleteByKeyAsync(keyValue);
        }



        public virtual int DeleteRange(IEnumerable<Entity> entities)
        {
            var keys = entities.Select(entity => entityDescriptor.key.GetValue(entity));
            return DeleteByKeys(keys);
        }
        public virtual Task<int> DeleteRangeAsync(IEnumerable<Entity> entities)
        {
            var keys = entities.Select(entity => entityDescriptor.key.GetValue(entity));
            return DeleteByKeysAsync<object>(keys);
        }


        public virtual int DeleteByKey(object keyValue) => DeleteByKeys(new[] { keyValue });
        public virtual Task<int> DeleteByKeyAsync(object keyValue) => DeleteByKeysAsync(new[] { keyValue });



        public virtual int DeleteByKeys<Key>(IEnumerable<Key> keys)
        {
            int count = DeleteByKeysWithoutSave(keys);
            Save();
            return count;
        }
        public virtual async Task<int> DeleteByKeysAsync<Key>(IEnumerable<Key> keys)
        {
            int count = DeleteByKeysWithoutSave(keys);
            await SaveAsync();
            return count;
        }

        #endregion


    }
}
