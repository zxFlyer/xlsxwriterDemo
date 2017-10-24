//
//  ViewController.m
//  ExcelDemo
//
//  Created by ZX on 2017/10/24.
//  Copyright © 2017年 ZX. All rights reserved.
//

#import "ViewController.h"
#import <xlsxwriter.h>

@interface ViewController ()

// 已加载到的行数
@property (nonatomic, assign) int rowNum;

/**
 “小计”所在的单元格
 */
@property (nonatomic, retain) NSMutableArray *sumArray;

@end

static lxw_workbook  *workbook;
static lxw_worksheet *worksheet;

static lxw_format *titleformat;// 各表格标题栏的格式
static lxw_format *leftcontentformat;// 最左侧一列内容的样式
static lxw_format *contentformat;// 内容的样式
static lxw_format *rightcontentformat;// 最右侧一列内容的样式
static lxw_format *leftsumformat;// 最左侧一列小计的样式
static lxw_format *sumformat;// 小计的样式
static lxw_format *rightsumformat;// 最右侧一列小计的样式

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    // Do any additional setup after loading the view, typically from a nib.
    
    UIButton *btn = [[UIButton alloc]initWithFrame:CGRectMake(100, 100, 100, 50)];
    btn.backgroundColor = [UIColor orangeColor];
    [btn setTitle:@"生成excel" forState:UIControlStateNormal];
    [btn addTarget:self action:@selector(btnClick) forControlEvents:UIControlEventTouchUpInside];
    [self.view addSubview:btn];
}
-(void)btnClick{
    NSMutableArray *trafficArray = [NSMutableArray array];
    NSMutableArray *mealsArray = [NSMutableArray array];
    NSMutableArray *travelArray = [NSMutableArray array];

    for (int i = 0; i < 5; i++) {
        NSDictionary *dic = @{
                              @"time": @"2017-10-19 11:15 至 11:15",
                              @"palce": @"广州-北京",
                              @"money": @"5"
                              };
        [trafficArray addObject:dic];
        [mealsArray addObject:dic];
        [travelArray addObject:dic];
    }
    NSDictionary *dataDic = @{
                              @"userinfo":@{
                                      @"username": @"阿轲",
                                      @"dateRange": @"报销日期范围：07年9月16日-07年10月15日"
                                      },
                              @"traffic": trafficArray,
                              @"meals": mealsArray,
                              @"travel": travelArray
                              };
    [self createXlsxFileWith:dataDic];
}

-(void)createXlsxFileWith:(NSDictionary *)dataDic{
    self.rowNum = 0;

    // 文件保存的路径
    NSString *documentPath = [NSSearchPathForDirectoriesInDomains(NSDocumentDirectory,NSUserDomainMask, YES) objectAtIndex:0];
    NSString *filename = [documentPath stringByAppendingPathComponent:@"c_demo.xlsx"];
    NSLog(@"filename_path:%@",filename);
    workbook  = workbook_new([filename UTF8String]);// 创建新xlsx文件，路径需要转成c字符串
    worksheet = workbook_add_worksheet(workbook, NULL);// 创建sheet
    [self setupFormat];
    
    [self createFormHeaderWithUserinfo:dataDic[@"userinfo"]];
    [self createTrafficForm:dataDic[@"traffic"]];
    [self createMealsForm:dataDic[@"meals"]];
    [self createTravelForm:dataDic[@"travel"]];
    [self creatOtherAndSumForm:dataDic[@"other"]];
    
    workbook_close(workbook);
}
// 单元格样式
-(void)setupFormat{
    titleformat = workbook_add_format(workbook);
    format_set_bold(titleformat);
    format_set_font_size(titleformat, 10);
    format_set_align(titleformat, LXW_ALIGN_CENTER);
    format_set_align(titleformat, LXW_ALIGN_VERTICAL_CENTER);//垂直居中
    format_set_border(titleformat, LXW_BORDER_MEDIUM);// 边框（四周）：中宽边框
    
    leftcontentformat = workbook_add_format(workbook);
    format_set_font_size(leftcontentformat, 10);
    format_set_left(leftcontentformat, LXW_BORDER_MEDIUM);// 左边框：中宽边框
    format_set_bottom(leftcontentformat, LXW_BORDER_DOUBLE);// 下边框：双线边框
    
    contentformat = workbook_add_format(workbook);
    format_set_font_size(contentformat, 10);
    format_set_left(contentformat, LXW_BORDER_DOUBLE);// 左边框：双线边框
    format_set_bottom(contentformat, LXW_BORDER_DOUBLE);// 下边框：双线边框
    format_set_right(contentformat, LXW_BORDER_DOUBLE);// 右边框：双线边框
    
    rightcontentformat = workbook_add_format(workbook);
    format_set_font_size(rightcontentformat, 10);
    format_set_bottom(rightcontentformat, LXW_BORDER_DOUBLE);// 下边框：双线边框
    format_set_right(rightcontentformat, LXW_BORDER_MEDIUM);// 右边框：中宽边框
    format_set_num_format(rightcontentformat, "￥#,##0.00");
    
    leftsumformat = workbook_add_format(workbook);
    format_set_font_size(leftsumformat, 10);
    format_set_left(leftsumformat, LXW_BORDER_MEDIUM);// 左边框：中宽边框
    format_set_bottom(leftsumformat, LXW_BORDER_MEDIUM);// 下边框：中宽边框
    
    sumformat = workbook_add_format(workbook);
    format_set_font_size(sumformat, 10);
    format_set_align(sumformat, LXW_ALIGN_RIGHT);// 右对齐
    format_set_left(sumformat, LXW_BORDER_DOUBLE);// 左边框：双线边框
    format_set_bottom(sumformat, LXW_BORDER_MEDIUM);// 下边框：中宽边框
    format_set_right(sumformat, LXW_BORDER_DOUBLE);// 右边框：双线边框
    
    rightsumformat = workbook_add_format(workbook);
    format_set_font_size(rightsumformat, 10);
    format_set_align(rightsumformat, LXW_ALIGN_RIGHT);// 右对齐
    format_set_bottom(rightsumformat, LXW_BORDER_MEDIUM);// 下边框：中宽边框
    format_set_right(rightsumformat, LXW_BORDER_MEDIUM);// 右边框：中宽边框
    format_set_num_format(rightsumformat, "￥#,##0.00");
}
// 整个文档的表头
-(void)createFormHeaderWithUserinfo:(NSDictionary *)userinfoDic{
    // 这个表格header标题格式
    lxw_format *headerFormat = workbook_add_format(workbook);
    format_set_font_size(headerFormat, 12);
    format_set_bold(headerFormat);
    format_set_align(headerFormat, LXW_ALIGN_CENTER);//水平居中
    format_set_align(headerFormat, LXW_ALIGN_VERTICAL_CENTER);//垂直居中
    
    
    // 姓名、报销日期格式
    lxw_format *nameFormat = workbook_add_format(workbook);
    format_set_font_size(nameFormat, 10);
    format_set_bold(nameFormat);
    
    // 设置列宽
    worksheet_set_column(worksheet, 1, 2, 30, NULL);// B、C两列宽度
    worksheet_set_column(worksheet, 3, 3, 25, NULL);// D列宽度
    
    worksheet_write_string(worksheet, self.rowNum, 2, "月报销申请表", headerFormat);
    worksheet_write_string(worksheet, ++self.rowNum, 0, "", NULL);//空白行
    NSString *username = [NSString stringWithFormat:@"申报人：%@", userinfoDic[@"username"]];
    const char *username_c = [username cStringUsingEncoding:NSUTF8StringEncoding];
    worksheet_write_string(worksheet, ++self.rowNum, 1, username_c, nameFormat);
    const char *dateRange_c = [userinfoDic[@"dateRange"] cStringUsingEncoding:NSUTF8StringEncoding];
    worksheet_write_string(worksheet, self.rowNum, 3, dateRange_c, nameFormat);
}
// 市内交通费表格
-(void)createTrafficForm:(NSArray *)dataArray{
    [self setupFormContent:dataArray titleString:@"市   内   交   通   费"];
}
// 市内餐费表格
-(void)createMealsForm:(NSArray *)dataArray{
    [self setupFormContent:dataArray titleString:@"市内餐费"];
}
// 差旅费表格
-(void)createTravelForm:(NSArray *)dataArray{
    [self setupFormContent:dataArray titleString:@"差旅费"];
}
// 其他费用、合计
-(void)creatOtherAndSumForm:(NSArray *)dataArray{
    [self setupFormContent:dataArray titleString:@"其他费用"];
}

-(void)setupFormContent:(NSArray *)dataArray titleString:(NSString *)titleString{
    worksheet_merge_range(worksheet, ++self.rowNum, 1, self.rowNum, 3, [titleString cStringUsingEncoding:NSUTF8StringEncoding], titleformat);
    if (![titleString isEqualToString:@"其他费用"]) {
        worksheet_write_string(worksheet, ++self.rowNum, 1, "日         期", titleformat);
        worksheet_write_string(worksheet, self.rowNum, 2, "来  往  地  点", titleformat);
        worksheet_write_string(worksheet, self.rowNum, 3, "金       额", titleformat);
    }
    
    int startRow = self.rowNum;
    for (int i = 0; i < dataArray.count; i++) {
        NSDictionary *dic = dataArray[i];
        worksheet_write_string(worksheet, ++self.rowNum, 1, [dic[@"time"] cStringUsingEncoding:NSUTF8StringEncoding], leftcontentformat);
        worksheet_write_string(worksheet, self.rowNum, 2,  [dic[@"place"] cStringUsingEncoding:NSUTF8StringEncoding], contentformat);
        worksheet_write_number(worksheet, self.rowNum, 3, [dic[@"money"] doubleValue], rightcontentformat);
    }
    // 空行
    worksheet_write_string(worksheet, ++self.rowNum, 1, "", leftcontentformat);
    worksheet_write_string(worksheet, self.rowNum, 2, "", contentformat);
    worksheet_write_number(worksheet, self.rowNum, 3, 0, rightcontentformat);
    
    int endRow = self.rowNum;
    NSString *sumFormula = [NSString stringWithFormat:@"=SUM(D%d:D%d)", startRow+1, endRow+1];
    worksheet_write_string(worksheet, ++self.rowNum, 1, "", leftsumformat);
    worksheet_write_string(worksheet, self.rowNum, 2, "小计：", sumformat);
    worksheet_write_formula(worksheet, self.rowNum, 3, [sumFormula cStringUsingEncoding:NSUTF8StringEncoding], rightsumformat);
    
    [self.sumArray addObject:@(self.rowNum+1)];// 记录小计金额单元格位置
    
    if ([titleString isEqualToString:@"其他费用"]) {
        [self sumTotalMoney];
    } else {
        worksheet_write_string(worksheet, ++self.rowNum, 0, "", NULL);// 空行
        worksheet_write_string(worksheet, ++self.rowNum, 0, "", NULL);// 空行
    }
}
-(void)sumTotalMoney{
    lxw_format *borderformat_alignleft = workbook_add_format(workbook);
    format_set_font_size(borderformat_alignleft, 10);
    format_set_border(borderformat_alignleft, LXW_BORDER_MEDIUM);//  边框（四周）：中宽边框
    worksheet_merge_range(worksheet, ++self.rowNum, 1, self.rowNum, 3, "备注：如有特殊说明请在此栏填写", borderformat_alignleft);
    
    lxw_format *borderformat_alignright = workbook_add_format(workbook);
    format_set_font_size(borderformat_alignright, 10);
    format_set_bold(borderformat_alignright);
    format_set_border(borderformat_alignright, LXW_BORDER_MEDIUM);//  边框（四周）：中宽边框
    format_set_align(borderformat_alignright, LXW_ALIGN_RIGHT);
    worksheet_write_string(worksheet, ++self.rowNum, 1, "", borderformat_alignleft);
    worksheet_write_string(worksheet, self.rowNum, 2, "合计", borderformat_alignright);
    
    lxw_format *totalmoneyformat = workbook_add_format(workbook);
    format_set_font_size(totalmoneyformat, 10);
    format_set_bold(totalmoneyformat);
    format_set_border(totalmoneyformat, LXW_BORDER_MEDIUM);//  边框（四周）：中宽边框
    format_set_align(totalmoneyformat, LXW_ALIGN_RIGHT);
    format_set_num_format(totalmoneyformat, "￥#,##0.00");
    NSString *sumStr = @"=D";
    for (int i = 0; i < self.sumArray.count; i++) {
        if (i < self.sumArray.count-1) {
            sumStr = [NSString stringWithFormat:@"%@%@+D", sumStr, self.sumArray[i]];
        } else {
            sumStr = [NSString stringWithFormat:@"%@%@", sumStr, self.sumArray[i]];
        }
    }
    NSLog(@"sumarray:%@", self.sumArray);
    worksheet_write_formula_num(worksheet, self.rowNum, 3, [sumStr cStringUsingEncoding:NSUTF8StringEncoding], totalmoneyformat, 0);
}
-(NSMutableArray *)sumArray{
    if (!_sumArray) {
        _sumArray = [NSMutableArray array];
    }
    return _sumArray;
}
@end
