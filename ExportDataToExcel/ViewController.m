//
//  ViewController.m
//  ExportDataToExcel
//
//  Created by ningcol on 10/13/16.
//  Copyright © 2016 ningcol. All rights reserved.
//
//
//   设置bitcode为no,other linker flag也要改为-lstdc++


#import "ViewController.h"

// **** 需要下载 LibXL.framework  *******
#include "LibXL/libxl.h"


@interface ViewController ()<UIDocumentInteractionControllerDelegate>
@property (nonatomic, strong) NSArray *nameArray;
@property (nonatomic, strong) NSArray *sexArray;
@property (nonatomic, strong) NSArray *schoolArray;
@property (nonatomic, strong) NSArray *phoneArray;
@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    
    self.nameArray = @[@"ningcol",@"张三",@"李四",@"王五"];
    self.sexArray = @[@"男",@"女",@"男",@"男"];
    self.schoolArray = @[@"北京大学",@"清华大学",@"复旦大学",@"家里蹲大学"];
    self.phoneArray = @[@"18345453452",@"13045453333",@"13845451112",@"180451111"];
    
    
    UIButton *btn = [UIButton buttonWithType:UIButtonTypeCustom];
    btn.frame = CGRectMake(120, 100, 80, 40);
    [btn setTitle:@"导出数据" forState:UIControlStateNormal];
    [btn setBackgroundColor:[UIColor redColor]];
    [btn addTarget:self action:@selector(clickBtn) forControlEvents:UIControlEventTouchUpInside];
    [self.view addSubview:btn];
}

-(void)clickBtn{
    

    BookHandle book = xlCreateBook(); // use xlCreateXMLBook() for working with xlsx files
    
    SheetHandle sheet = xlBookAddSheet(book, "Sheet1", NULL);
    //第一个参数代表插入哪个表，第二个是第几行（默认从0开始），第三个是第几列（默认从0开始）
    xlSheetWriteStr(sheet, 1, 0, "姓名", 0);
    xlSheetWriteStr(sheet, 1, 1, "性别", 0);
    xlSheetWriteStr(sheet, 1, 2, "学校", 0);
    xlSheetWriteStr(sheet, 1, 3, "电话", 0);
    
    
    for (int i = 0; i < self.nameArray.count; i++) {
        const char *name_c = [self.nameArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i+2, 0,name_c, 0);
        
    }
    for (int i = 0; i < self.sexArray.count; i++) {
        const char *sex_c = [self.sexArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i+2, 1,sex_c, 0);
        
    }
    for (int i = 0; i < self.schoolArray.count; i++) {
        const char *school_c = [self.schoolArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i+2, 2,school_c, 0);
        
    }
    for (int i = 0; i < self.phoneArray.count; i++) {
        const char *phone_c = [self.phoneArray[i] cStringUsingEncoding:NSUTF8StringEncoding];
        xlSheetWriteStr(sheet, i+2, 3,phone_c, 0);
        
    }
    
    
    NSString *documentPath =
    [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory,NSUserDomainMask, YES) objectAtIndex:0];
    NSString *fname = [@"data" stringByAppendingString:@".xls"];
    NSString *filename = [documentPath stringByAppendingPathComponent:fname];
    // 输出路径
    NSLog(@"filepath:%@",filename);
    
    xlBookSave(book, [filename UTF8String]);
    
    xlBookRelease(book);
    
    // 分享出去
    UIDocumentInteractionController *docu = [UIDocumentInteractionController interactionControllerWithURL:[NSURL fileURLWithPath:filename]];
    
    docu.delegate = self;
    CGRect rect = CGRectMake(0, 0, 320, 300);
    
    [docu presentOpenInMenuFromRect:rect inView:self.view animated:YES];
    [docu presentPreviewAnimated:YES];

    
}



@end
