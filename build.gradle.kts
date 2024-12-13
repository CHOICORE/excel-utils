plugins {
    java
}

group = "me.choicore.samples"
version = "0.0.1-SNAPSHOT"

java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(21)
    }
}

configurations {
    compileOnly {
        extendsFrom(configurations.annotationProcessor.get())
    }
}

repositories {
    mavenCentral()
}

dependencies {
    annotationProcessor("org.projectlombok:lombok:1.18.36")
    compileOnly("org.projectlombok:lombok:1.18.36")
    implementation("org.apache.poi:poi-ooxml:5.3.0")
    implementation("org.apache.poi:poi:5.3.0")
}
