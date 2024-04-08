# Excel Plus

基于 Apache POI 库，具体使用请参考 `example` 包

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
   
    ...

   }
   ```
2. 配置依赖
   ```
   dependencies {
    implementation('com.ericyl.excel:excel-plus:0.1.0-SNAPSHOT')
    implementation("org.projectlombok:lombok:${lombokVersion}")
    annotationProcessor("org.projectlombok:lombok:${lombokVersion}")
    implementation("org.apache.commons:commons-lang3:${lang3Version}")
    implementation("commons-collections:commons-collections:${collectionsVersion}")
    implementation("org.apache.poi:poi:${poiVersion}")
    implementation("org.apache.poi:poi-ooxml:${poiVersion}")
   }
   ```
