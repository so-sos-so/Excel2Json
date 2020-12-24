using System.Collections.Generic;

public class Troops
{
	public int ID { get; set; }// ID
	public String Name { get; set; }// 名字
	public String icon { get; set; }// 图标
	public float maxHealth { get; set; }// 初始最大血量
	public float attack { get; set; }// 攻击
	public float defense { get; set; }// 防御
	public float attackSpeed { get; set; }// 攻击速度
	public float moveSpeed { get; set; }// 移动速度
	public float attackRange { get; set; }// 攻击范围
	public int fightingCapacity { get; set; }// 战斗力
	public int occupied_X { get; set; }// 占位长
	public int occupied_Y { get; set; }// 占位宽
	public String modelName { get; set; }// 模型名字
	public float strength { get; set; }// 英雄属性:力量
	public float intelligence { get; set; }// 英雄属性:智力
	public float commander { get; set; }// 英雄属性:统帅
	public float charm { get; set; }// 英雄属性:魅力
	public int configuration_X { get; set; }// 英雄属性:配置区域长
	public int configuration_Y { get; set; }// 英雄属性:配置区域宽
	public float damage { get; set; }// 兵种属性:伤害
	public float avoidInjury { get; set; }// 兵种属性:免伤
	public float criticalStrike { get; set; }// 兵种属性:暴击
	public float knockRating { get; set; }// 兵种属性:抗暴
	public float parry { get; set; }// 兵种属性:格挡
	public float atypic { get; set; }// 兵种属性:破格
	public String soldiersModelName { get; set; }// 兵种属性:士兵模型名字
	public int soldiersRow { get; set; }// 兵种属性:士兵排数
	public int soldiersColumn { get; set; }// 兵种属性:士兵列数

}