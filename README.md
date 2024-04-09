# Excel Plus

基于 Apache POI 库，可配置数据处理器

* Excel 读取
  * 单类读取
    * 通过 Excel 单元格坐标读取数据 `@ExcelReader(value = 'A5')`
  * 列表读取
    * 通过 Excel 下标 `@ExcelReader(index = 0)`
    * 【支持多表头】通过 Excel 表头 `@ExcelReader(name = {'表头1','表头2'})`
    
* Excel 生成
  * 【支持分页】简单列表可通过配置 `@ExcelWriter(name = '名称')` 生成 Excel
  * 复杂 Excel 需自定义 `ExcelTable` 的方式进行处理

具体使用方法请参考 `com.ericyl.excel.example` 包中的示例

## 怎么使用

### Gradle
1. 配置 maven 库
   ```
   repositories {
   
   //使用Github Packages
   //    maven {
   //        url = uri("https://maven.pkg.github.com/ericyl-studio/excel-plus")
   //        credentials {
   //            username = "GITHUB_USERNAME"
   //            password = "GITHUB_TOKEN"
   //        }
   //    }
   
     maven {
       url = uri("https://oss.sonatype.org/content/repositories/snapshots/")
     }
   
    //...

   }
   ```
2. 配置依赖
   ```
   dependencies {
    implementation('com.ericyl.excel:excel-plus:0.1.1-SNAPSHOT')
    implementation("org.projectlombok:lombok:${lombokVersion}")
    annotationProcessor("org.projectlombok:lombok:${lombokVersion}")
    implementation("org.apache.commons:commons-lang3:${lang3Version}")
    implementation("org.apache.commons:commons-collections4:${collectionsVersion}")
    implementation("org.apache.poi:poi:${poiVersion}")
    implementation("org.apache.poi:poi-ooxml:${poiVersion}")
   }
   ```

## 使用到的类库
1. [Lombok](https://github.com/projectlombok/lombok)
2. [Apache commons](https://commons.apache.org)
3. [Apache POI](https://poi.apache.org)