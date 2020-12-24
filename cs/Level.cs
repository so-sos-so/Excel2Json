using System.Collections.Generic;

public class Level
{
	public int ID { get; set; }// #
	public int level { get; set; }// 关卡
	public int hero { get; set; }// 关卡英雄
	public int occupied { get; set; }// 占位(1-3)
	public int corpTroops1 { get; set; }// 兵种部队1
	public List<string> crop1Occuopied { get; set; }// 兵种1占位
	public int corpTroops2 { get; set; }// 兵种部队2
	public List<string> crop2Occuopied { get; set; }// 兵种2占位
	public int corpTroops3 { get; set; }// 兵种部队3
	public List<string> crop3Occuopied { get; set; }// 兵种3占位

}