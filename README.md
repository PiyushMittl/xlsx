# Apache POI – Reading and Writing Excel file in Java

## The EndPoint Abstract class:
First class that generalizes both producers and consumers as ‘endpoints’ of a queue. Whether you are a producer or a consumer, the code to connect to a queue remains the same therefore we have a common class EndPoint for that.

## The Producer:
The producer class is what is responsible for writing a message onto a queue. We are using Apache Commons Lang to convert a Serializable java object to a byte array. The required maven dependency for commons lang is given in below Maven section.

## The Consumer:
The consumer, which can be run as a thread, has callback functions for various events, most important of which is the availability of a new message.



### Setup


provide following:

```java
```
 

### Maven

The maven dependency for the java client is given below.

```xml
	<dependency>
        <groupId>com.rabbitmq</groupId>
        <artifactId>amqp-client</artifactId>
        <version>3.0.4</version>
	</dependency> 
```

The maven dependency for commons lang is 

```xml
	<dependency>
		<groupId>commons-lang</groupId>
		<artifactId>commons-lang</artifactId>
		<version>2.6</version>
	</dependency>
```



the following code in main is used to run the sending mail.

```java
```


references:
https://www.mkyong.com/java/apache-poi-reading-and-writing-excel-file-in-java/
http://howtodoinjava.com/apache-commons/readingwriting-excel-files-in-java-poi-tutorial/