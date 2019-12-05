import scrapy.cmdline

if __name__ == "__main__":
    print("HELLO, SPIDER!")
    scrapy.cmdline.execute(
        "scrapy crawl place -a country=Sri-Lanka -a interests=Airport".split())
