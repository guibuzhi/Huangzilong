using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;


using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
//using LY;
using System.IO;
//using System.Windows.Forms;
using System.Data.SqlClient;
//using LY.MsSqlHelper;
//using ICSharpCode.SharpZipLib.Zip;
using System.Threading;
//using external;
using ClassLibrary1;
using System.Data;
using MySql.Data.MySqlClient;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;

namespace Huangzilong
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void 转子冲片系列(object sender, EventArgs e)
        {
            //if (Session["path"] == null)
            //{
            //    Response.Write("<script>alert('您还没有进行方案设计，请先点击方案设计按钮生成设计方案！')</script>");
            //}
            //else
            //{ }

            foreach (ListItem a in ListBox2.Items)
            {
                string code;
                if (a.Selected == true)
                {
                    SldWorks.SldWorks swapp = new SldWorks.SldWorks();
                    swapp.Visible = true;
                    swapp.FrameState = 1;

                    code = a.Value;
                    if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "Roundbottomslot_roundhole_Squarehole_middlehole", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Roundbottomslot_roundhole_Squarehole_middlehole(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "Roundbottomslot_Localarrayhole", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Roundbottomslot_Localarrayhole(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "Three_rectangular_slots", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Three_rectangular_slots(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "Flat_bottom_slots", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Flat_bottom_slots(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "flatbottom_smallroundbottom_slot", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Flatbottom_smallroundbottom_slot(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "v_shaped_hole", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().V_shaped_hole(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "half_ring_hole", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().half_ring_hole(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "Parallelogram_slot", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Parallelogram_slot(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "bend_slot", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().bend_slot(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "double_rectangle_slot", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().double_rectangle_slot(code);
                    else if (new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().Query_database("rotor_lamination", "single_rectangular_slot", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().single_rectangular_slot(code);
                    

                }
            }

            //{ LY.JsHelper.AlertAndRedirect("绘图完成", "Default1.aspx"); }

           
        }


        protected void EQ_214_2431(object sender, EventArgs e)
        {
            SldWorks.SldWorks swapp = new SldWorks.SldWorks();
            swapp.Visible = true;
            swapp.FrameState = 1;

            foreach (ListItem a in ListBox6.Items)
            {
                if (a.Selected == true)
                {
                    //new ClassLibrary1.Class1().初始设置(2);
                    //Program program = new Program();
                    //program.KillProcess("SLDWORKS");
                    //Thread.Sleep(2000);
                    

                    string filename = a.Value;
                    //program.Drawing_ly25_026(filename);

                    if (filename == "rotor_lamination")
                        ////huangzilong
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().rotor_lamination();
                    if (filename == "damping_plate")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().damping_plate();
                    if (filename == "connecting_shaft_piece")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().connecting_shaft_piece();
                    if (filename == "Exciter_rotor_pressing_ring")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().Exciter_rotor_pressing_ring();
                    if (filename == "support_block_screw")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().support_block_screw();
                    if (filename == "Support_block_2004")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().Support_block_2004();
                    if (filename == "Support_block_2005")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().Support_block_2005();
                    if (filename == "Support_block_assembly")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().Support_block_assembly();

                    if (filename == "Fan_parts")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().fan_parts_series_one("8LY.435.002", "默认");
                    if (filename == "Rotor_core")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().Rotor_core();
                    if (filename == "线圈")
                        new ClassLibrary1.Huangzilong.Run_huangzilong.EQ_214_2431.Parts_drawing().线圈();

                    

                }
            }
        }

        protected void fan(object sender, EventArgs e)
        {
            foreach (ListItem a in ListBox1.Items)
            {
                string code;
                if (a.Selected == true)
                {
                    SldWorks.SldWorks swapp = new SldWorks.SldWorks();
                    swapp.Visible = true;
                    swapp.FrameState = 1;

                    code = a.Value;
                    if (new ClassLibrary1.Huangzilong.Module_huangzilong().Query_database("fan", "Outside_fan", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.fan().外风扇_8LY_435_001(code);
                    if (new ClassLibrary1.Huangzilong.Module_huangzilong().Query_database("fan", "fan_parts_series_one", code))
                        new ClassLibrary1.Huangzilong.Run_huangzilong.fan().fan_parts_series_one(code,"默认");


                }
            }
            //foreach (ListItem a in ListBox1.Items)
            //{
            //    
            //    if (a.Selected == true)
            //    {



            //    }
            //}


        }
        protected void Y2(object sender, EventArgs e)
        {
            ClassLibrary1.Huangzilong.Module_huangzilong Module_huangzilong = new ClassLibrary1.Huangzilong.Module_huangzilong();
            ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination rotor_lamination = new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination();
            ClassLibrary1.Huangzilong.Run_huangzilong.rotor_shaft rotor_shaft = new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_shaft();
            string alihost, localhost, alidatabase, localdatabase, alitable_name, localtable_name;
            alihost = "localhost";
            localhost = "localhost";
            alidatabase = "baserequires";
            alitable_name = "optimal_solution_copy2";
            //int ID = 720;//355
            //int ID = 721;//180
            //int ID = 722;//200
            //int ID = 727;//200
            //int ID = 726;//180
            //int ID = 728;//280
            int ID = 720;//355
            //int ID = 708;//250
            Module_huangzilong.inputid(ID);
            Module_huangzilong.set_waittime(localhost);

            //double[] points = { 1, 2, 3, 4, 5, 6, 7 };
            //foreach (double point in points)
            //{
            //rotor_lamination.sheet2("code" + point.ToString(), localhost);
            //}

            //MessageBox.Show("55");
            //double[] points = { 1, 2, 3, 4, 5, 6, 7 };
            //foreach (double point in points)
            //{
            //rotor_shaft.sheet1("code1", localhost, "006");//5阶
            //rotor_shaft.sheet1("code2", localhost, "007");//9阶
            //rotor_shaft.sheet1("code3", localhost, "008");//7阶
            //rotor_shaft.sheet1("code4", localhost, "008");
            //rotor_shaft.sheet1("code5", localhost, "006");
            //rotor_shaft.sheet1("code6", localhost, "007");
            //rotor_shaft.sheet1("code7", localhost, "008");
            //rotor_shaft.sheet1("code8", localhost, "008");
            rotor_shaft.sheet1("code9", localhost, "009");//10阶












            //}
            //6个,第一份工作做了7年
            //


            MessageBox.Show("55");
        }

        protected void 同步发电机总装配新(object sender, EventArgs e)
        {
            ClassLibrary1.Huangzilong.Module_huangzilong Module_huangzilong = new ClassLibrary1.Huangzilong.Module_huangzilong();
            ClassLibrary1.Huangzilong.Run_huangzilong Run_huangzilong = new ClassLibrary1.Huangzilong.Run_huangzilong();

            string alihost, localhost, alidatabase, localdatabase, alitable_name, localtable_name;
            alihost = "localhost";
            localhost = "localhost";
            alidatabase = "baserequires";
            alitable_name = "optimal_solution_copy2";
            //int ID = 720;//355
            //int ID = 721;//180
            //int ID = 722;//200
            //int ID = 727;//200
            //int ID = 726;//180
            //int ID = 728;//280
            int ID = 720;//355
            //int ID = 708;//250
            Module_huangzilong.inputid(ID);
            Module_huangzilong.set_waittime(localhost);
            ClassLibrary1.Module_wangbo Module_wangbo = new ClassLibrary1.Module_wangbo();
            Module_wangbo.inputid(ID);

            //阻尼棒
            string[] Damping_rod_code;
            Damping_rod_code = Module_huangzilong.Damping_rod(localhost, "Damping_rod", "damping_rod", alihost, alidatabase, alitable_name);
            if ("repeat" == Damping_rod_code[1] && Module_huangzilong.check_file(localhost, "Damping_rod", "damping_rod", Damping_rod_code[0])) ;
            else
            {
                new ClassLibrary1.Huangzilong.Run_huangzilong.Damping_rod().Damping_rod_(Damping_rod_code[0], localhost);
            }

            //转子铁芯
            string[] rotor_core_code;
            rotor_core_code = Module_huangzilong.single_rectangular_slot_rotor_core(localhost, "rotor_lamination", "single_rectangular_slot", alihost, alidatabase, alitable_name);
            if ("repeat" == rotor_core_code[1] && Module_huangzilong.check_file(localhost, "rotor_lamination", "single_rectangular_slot", rotor_core_code[0])) ;
            else
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().single_rectangular_slot_(rotor_core_code[0], "rotor_lamination", "single_rectangular_slot", dz: localhost);

            //阻尼板
            string[] Damping_plate_code;
            Damping_plate_code = Module_huangzilong.single_rectangular_slot_Damping_plate(localhost, "Damping_plate", "single_rectangular_slot", alihost, alidatabase, alitable_name);
            if ("repeat" == Damping_plate_code[1] && Module_huangzilong.check_file(localhost, "Damping_plate", "single_rectangular_slot", Damping_plate_code[0])) ;
            else
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().single_rectangular_slot_Damping_plate(Damping_plate_code[0], "Damping_plate", "single_rectangular_slot", localhost, "红铜-钴-铍-合金，UNS C17500");

            //转子铁芯装配
            string[] rotor_core_assembly_code;
            rotor_core_assembly_code = Module_huangzilong.single_rectangular_slot_rotor_core_assembly(localhost, "rotor_core_assembly", "single_rectangular_slot_rotor_core", Damping_rod_code[0], rotor_core_code[0], Damping_plate_code[0]);
            if ("repeat" == rotor_core_assembly_code[1] && Module_huangzilong.check_file(localhost, "rotor_core_assembly", "single_rectangular_slot_rotor_core", rotor_core_assembly_code[0])) ;
            else
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_core_assembly().rotor_core_assembly(rotor_core_assembly_code[0], localhost);

            //查询机座号
            string housenumber;
            housenumber = Module_huangzilong.Read_house_number(alihost, alidatabase, alitable_name);
            //转轴
            string[] Shaft_code;
            string Shaft_table_name = "house_number_" + housenumber + "_shaft";
            Shaft_code = Module_huangzilong.Shaft_Update(localhost, "shaft", Shaft_table_name, alihost, alidatabase, alitable_name);
            if ("repeat" == Shaft_code[1] && Module_huangzilong.check_file(localhost, "shaft", Shaft_table_name, Shaft_code[0])) ;
            else if (housenumber == "180")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().轴180(localhost, "Shaft", "house_number_180_shaft", Shaft_code[0]);
            //new ClassLibrary1.fubin.Module_fubin.all.open_initialization

            else if (housenumber == "200")
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().house_number_200_shaft(Shaft_code[0], "Shaft", "house_number_200_shaft", localhost);
            else if (housenumber == "250")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().轴250(localhost, "Shaft", "house_number_250_shaft", Shaft_code[0]);
            //new ClassLibrary1.rotor_shaft.Run.Shaft().house_number_250_shaft(Shaft_code[0]);
            else if (housenumber == "280")
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().house_number_280_shaft(Shaft_code[0], "Shaft", "house_number_280_shaft", localhost);
            else if (housenumber == "355")
                new ClassLibrary1.zhangweibin.Run_huangzilong.Y2_450().转轴355尺寸驱动(localhost, "shaft", "house_number_355_shaft", Shaft_code[0]);
            //MessageBox.Show("444");
            //new ClassLibrary1.rotor_shaft.Run.Shaft().house_number_355_shaft(Shaft_code[0]);





            //风扇
            string[] fan_code;
            fan_code = Module_huangzilong.fan(housenumber);

            //shaft_sleeve‘联轴器‘励磁机转子
            string[] shaft_sleeve_code;
            shaft_sleeve_code = Module_huangzilong.shaft_sleeve(housenumber);

            //balance_ring‘平衡环
            string[] balance_ring_code;
            balance_ring_code = Module_huangzilong.balance_ring(housenumber);

            //coupling_piece’联轴器片
            string[] coupling_piece_code;
            coupling_piece_code = Module_huangzilong.coupling_piece(housenumber);


            //转子装配
            string[] rotor_assembly_code;
            rotor_assembly_code = Module_huangzilong.single_rectangular_slot_rotor_assembly(localhost, "rotor_assembly", "single_rectangular_slot_rotor", rotor_core_assembly_code[0], Shaft_code[0], fan_code[0], fan_code[1], shaft_sleeve_code[1], housenumber, shaft_sleeve_code[0], balance_ring_code[0], coupling_piece_code[0]);
            if ("repeat" == rotor_assembly_code[1] && Module_huangzilong.check_file(localhost, "rotor_assembly", "single_rectangular_slot_rotor", rotor_assembly_code[0])) ;
            else if (housenumber == "250")
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_assembly().rotor_250_assembly(rotor_assembly_code[0]);
            else if (housenumber == "355")
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_assembly().rotor_355_assembly(rotor_assembly_code[0]);
            else if (housenumber == "280")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().转子280装配(rotor_assembly_code[0]);
            else if (housenumber == "200")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().转子200装配(rotor_assembly_code[0]);
            else if (housenumber == "180")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().转子180装配(rotor_assembly_code[0]);



            //MessageBox.Show("111");
            string[] sl, ring, clasp, stator_core_assembly;

            sl = Module_wangbo.stator_lamination_mysql(localhost, alihost, alitable_name);
            if (sl[1] == "repeat" && Module_huangzilong.check_file(localhost, "stator_lamination", "all_domeslot_lamination", sl[0], "part_save_address")) ;
            else
            {
                //new ClassLibrary1.Run_wangbo.Stator_lamination().all_domeslot_lamination(sl[0], localhost);
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().定子冲片(localhost, "stator_lamination", "all_domeslot_lamination", sl[0]);

            }

            ring = Module_wangbo.stator_pressing_ring_mysql(localhost, alitable_name);
            if (ring[1] == "repeat" && Module_huangzilong.check_file(localhost, "stator_pressing_ring", "stator_pressing_ring", ring[0], "part_save_address")) ;
            else
            {
                //new ClassLibrary1.Run_wangbo.LY_synchronous_motor.Stator().stator_pressing_ring(ring[0], localhost);
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().定子端环(localhost, "stator_pressing_ring", "stator_pressing_ring", ring[0]);

            }

            clasp = Module_wangbo.clasp_mysql(localhost, sl[0], ring[0]);
            if (clasp[1] == "repeat" && Module_huangzilong.check_file(localhost, "clasp", "clasp", clasp[0], "part_save_address")) ;
            else
            {
                //new ClassLibrary1.Run_wangbo.LY_synchronous_motor.Stator().clasp(clasp[0], localhost);
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().扣片(localhost, "clasp", "clasp", clasp[0]);

            }

            stator_core_assembly = Module_wangbo.stator_core_assembly_mysql(localhost, alihost, alitable_name);
            if (stator_core_assembly[1] == "repeat" && Module_huangzilong.check_file(localhost, "stator_core_assembly", "stator_core_assembly", stator_core_assembly[0], "part_save_address")) ;
            else
            {
                new ClassLibrary1.Run_wangbo.LY_synchronous_motor.Stator().stator_core_assembly(stator_core_assembly[0], localhost, clasp[0]);
            }


            //MessageBox.Show("定子结束");

            //励磁机定子装配
            string[] Exciter_stator_core_assembly;
            Exciter_stator_core_assembly = Module_huangzilong.Exciter_stator_core_assembly(housenumber);



            //前端盖
            string[] front_end_cover_code;
            front_end_cover_code = Module_huangzilong.front_end_cover(housenumber);



            string[] main_frame;
            main_frame = Module_huangzilong.main_frame(housenumber);
            //MessageBox.Show("定子结束");




            string General_assembly_address, General_pack_address, pdf_pack_address;
            General_assembly_address = "E:\\works\\generator_parts_library\\generator_assembly\\";
            General_pack_address = "C:\\Users\\Administrator\\Desktop\\pack\\";
            pdf_pack_address = "C:\\Users\\Administrator\\Desktop\\pdf\\";

            //总装配
            string[] generator_assembly_code;


            generator_assembly_code = new ClassLibrary1.Huangzilong.Module_huangzilong().generator_assembly(localhost, "generator_assembly", "generator_assembly", rotor_assembly_code[0], stator_core_assembly[0], Exciter_stator_core_assembly[0], front_end_cover_code[0], main_frame[0], housenumber);

            //MessageBox.Show(new ClassLibrary1.Huangzilong.Module_huangzilong().check_file(localhost, "generator_assembly", "generator_assembly", generator_assembly_code[0]).ToString());
            //MessageBox.Show(generator_assembly_code[1]);
            Module_huangzilong.修改筒体长度(Shaft_code[0], housenumber, localhost);
            Module_huangzilong.修改筋长度(Shaft_code[0], housenumber, localhost);
            if ("repeat" == generator_assembly_code[1] && new ClassLibrary1.Huangzilong.Module_huangzilong().check_file(localhost, "generator_assembly", "generator_assembly", generator_assembly_code[0])) ;

            else if (housenumber == "355")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配355(generator_assembly_code[0]);
            else if (housenumber == "250")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配250(generator_assembly_code[0]);
            else if (housenumber == "280")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配280(generator_assembly_code[0]);
            else if (housenumber == "200")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配200(generator_assembly_code[0]);
            else if (housenumber == "180")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配180(generator_assembly_code[0]);
            //MessageBox.Show("555");


            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_and_go(General_assembly_address + generator_assembly_code[0], General_pack_address);



            //定子冲片pdf
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + sl[0] + "_LY", pdf_pack_address + "LY_" + sl[0] + "_LY");
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + sl[0] + "_LY", pdf_pack_address + "LY_" + sl[0] + "_LY", 3);

            //前端盖pdf
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + front_end_cover_code[0] + "_LY", pdf_pack_address + "LY_" + front_end_cover_code[0] + "_LY");
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + front_end_cover_code[0] + "_LY", pdf_pack_address + "LY_" + front_end_cover_code[0] + "_LY", 3);

            //风扇pdf
            if (housenumber == "355" && housenumber == "280" && housenumber == "200")
            {
                new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + fan_code[0] + "_LY", pdf_pack_address + "LY_" + fan_code[0] + "_LY", 2);
                new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + fan_code[0] + "_LY", pdf_pack_address + "LY_" + fan_code[0] + "_LY", 3);
            }
            else
            {
                new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + fan_code[0] + "_LY", pdf_pack_address + "LY_" + fan_code[0] + "_LY", 1);
                new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + fan_code[0] + "_LY", pdf_pack_address + "LY_" + fan_code[0] + "_LY", 3);
            }

            //后端盖pdf
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + Exciter_stator_core_assembly[2] + "_LY", pdf_pack_address + "LY_" + Exciter_stator_core_assembly[2] + "_LY");
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + Exciter_stator_core_assembly[2] + "_LY", pdf_pack_address + "LY_" + Exciter_stator_core_assembly[2] + "_LY", 3);

            //转子冲片pdf
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + rotor_core_code[0] + "_LY", pdf_pack_address + "LY_" + rotor_core_code[0] + "_LY");
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + rotor_core_code[0] + "_LY", pdf_pack_address + "LY_" + rotor_core_code[0] + "_LY", 3);

            //转轴
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + Shaft_code[0] + "_LY", pdf_pack_address + "LY_" + Shaft_code[0] + "_LY");



        }


        protected void 阻尼板(object sender, EventArgs e)
        {
            string[] code;
            code = new ClassLibrary1.Huangzilong.Module_huangzilong().single_rectangular_slot_Damping_plate("localhost", "Damping_plate", "single_rectangular_slot", "106.15.236.225", "baserequires", "optimize_rotor_parameter");
            if ("repeat" == code[1]) ;
            else
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().single_rectangular_slot(code[0], "damping_plate", material: "红铜-钴-铍-合金，UNS C17500");
            ////注意使用
            ////转子
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "Damping_rod", "damping_rod");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "rotor_lamination", "single_rectangular_slot");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "Damping_plate", "single_rectangular_slot");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "shaft", "house_number_250_shaft");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "shaft", "house_number_180_shaft");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "shaft", "house_number_280_shaft");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "shaft", "house_number_200_shaft");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "shaft", "house_number_355_shaft");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "rotor_core_assembly", "single_rectangular_slot_rotor_core");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "rotor_assembly", "single_rectangular_slot_rotor");
            //定子
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "stator_lamination", "all_domeslot_lamination");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "stator_pressing_ring", "stator_pressing_ring");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "clasp", "clasp");
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "stator_core_assembly", "stator_core_assembly");
            ////总装
            //new ClassLibrary1.Huangzilong.Module_huangzilong().reset(localhost, "generator_assembly", "generator_assembly");
        }

        protected void 同步发电机转子铁芯(object sender, EventArgs e)
        {
            ClassLibrary1.Huangzilong.Module_huangzilong Module_huangzilong = new ClassLibrary1.Huangzilong.Module_huangzilong();
            //Module_huangzilong.Modify_feature();

        }





        protected void 总装配(object sender, EventArgs e)
        {




            string alihost, localhost, alidatabase, localdatabase, alitable_name, localtable_name;
            alihost = "localhost";
            localhost = "localhost";
            //localhost = "106.15.236.225";
            alidatabase = "baserequires";
            //alitable_name = "optimal_solution";
            alitable_name = "optimal_solution_copy2";
            //int ID = 727;//200
            //int ID = 726;//180
            int ID = 728;//280
            //int ID = 720;//355
            //int ID = 708;//250

            ClassLibrary1.Huangzilong.Module_huangzilong Module_huangzilong = new ClassLibrary1.Huangzilong.Module_huangzilong();
            ClassLibrary1.Module_wangbo Module_wangbo = new ClassLibrary1.Module_wangbo();

            Module_huangzilong.inputid(ID);
            Module_huangzilong.set_waittime(localhost);
            Module_wangbo.inputid(ID);
            // MessageBox.Show(Module.SS());


            //阻尼棒
            string[] Damping_rod_code;
            Damping_rod_code = Module_huangzilong.Damping_rod(localhost, "Damping_rod", "damping_rod", alihost, alidatabase, alitable_name);
            if ("repeat" == Damping_rod_code[1] && Module_huangzilong.check_file(localhost, "Damping_rod", "damping_rod", Damping_rod_code[0]));
            else
            { new ClassLibrary1.Huangzilong.Run_huangzilong.Damping_rod().Damping_rod(Damping_rod_code[0], localhost);
            }


            //转子铁芯
            string[] rotor_core_code;
            rotor_core_code = Module_huangzilong.single_rectangular_slot_rotor_core(localhost, "rotor_lamination", "single_rectangular_slot", alihost, alidatabase, alitable_name);
            if ("repeat" == rotor_core_code[1] && Module_huangzilong.check_file(localhost, "rotor_lamination", "single_rectangular_slot", rotor_core_code[0])) ;
            else
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().single_rectangular_slot(rotor_core_code[0], dz: localhost);
            //MessageBox.Show("666");



            //阻尼板
            string[] Damping_plate_code;
            Damping_plate_code = Module_huangzilong.single_rectangular_slot_Damping_plate(localhost, "Damping_plate", "single_rectangular_slot", alihost, alidatabase, alitable_name);
            //MessageBox.Show(Damping_plate_code[0]);
            if ("repeat" == Damping_plate_code[1] && Module_huangzilong.check_file(localhost, "Damping_plate", "single_rectangular_slot", Damping_plate_code[0])) ;
            else
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_lamination().single_rectangular_slot(Damping_plate_code[0], "damping_plate", material: "红铜-钴-铍-合金，UNS C17500", dz: localhost);
            //MessageBox.Show("666");



            //转子铁芯装配
            string[] rotor_core_assembly_code;
            rotor_core_assembly_code = Module_huangzilong.single_rectangular_slot_rotor_core_assembly(localhost, "rotor_core_assembly", "single_rectangular_slot_rotor_core", Damping_rod_code[0], rotor_core_code[0], Damping_plate_code[0]);
            if ("repeat" == rotor_core_assembly_code[1] && Module_huangzilong.check_file(localhost, "rotor_core_assembly", "single_rectangular_slot_rotor_core", rotor_core_assembly_code[0])) ;
            else
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_core_assembly().rotor_core_assembly(rotor_core_assembly_code[0], localhost);



            //查询机座号
            string housenumber;
            housenumber = Module_huangzilong.Read_house_number(alihost, alidatabase, alitable_name);
            //转轴
            string[] Shaft_code;
            string Shaft_table_name = "house_number_" + housenumber + "_shaft";
            Shaft_code = Module_huangzilong.Shaft_Update(localhost, "shaft", Shaft_table_name, alihost, alidatabase, alitable_name);
            if ("repeat" == Shaft_code[1] && Module_huangzilong.check_file(localhost, "shaft", Shaft_table_name, Shaft_code[0])) ;
            else if (housenumber == "180")
                new ClassLibrary1.rotor_shaft.Run.Shaft().house_number_180_shaft(Shaft_code[0]);
            else if (housenumber == "200")
                new ClassLibrary1.rotor_shaft.Run.Shaft().house_number_200_shaft(Shaft_code[0]);
            else if (housenumber == "250")
                new ClassLibrary1.rotor_shaft.Run.Shaft().house_number_250_shaft(Shaft_code[0]);
            else if (housenumber == "280")
                new ClassLibrary1.rotor_shaft.Run.Shaft().house_number_280_shaft(Shaft_code[0]);
            else if (housenumber == "355")
                new ClassLibrary1.rotor_shaft.Run.Shaft().house_number_355_shaft(Shaft_code[0]);


            

            //风扇
            string[] fan_code;
            fan_code = Module_huangzilong.fan(housenumber);

            //shaft_sleeve‘联轴器‘励磁机转子
            string[] shaft_sleeve_code;
            shaft_sleeve_code = Module_huangzilong.shaft_sleeve(housenumber);

            //balance_ring‘平衡环
            string[] balance_ring_code;
            balance_ring_code = Module_huangzilong.balance_ring(housenumber);

            //coupling_piece’联轴器片
            string[] coupling_piece_code;
            coupling_piece_code = Module_huangzilong.coupling_piece(housenumber);


            //转子装配
            string[] rotor_assembly_code;
            rotor_assembly_code = Module_huangzilong.single_rectangular_slot_rotor_assembly(localhost, "rotor_assembly", "single_rectangular_slot_rotor", rotor_core_assembly_code[0], Shaft_code[0], fan_code[0], fan_code[1], shaft_sleeve_code[1], housenumber, shaft_sleeve_code[0],  balance_ring_code[0], coupling_piece_code[0]);
            if ("repeat" == rotor_assembly_code[1] && Module_huangzilong.check_file(localhost, "rotor_assembly", "single_rectangular_slot_rotor", rotor_assembly_code[0])) ;
            else if (housenumber == "250")
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_assembly().rotor_250_assembly(rotor_assembly_code[0]);
            else if (housenumber == "355")
                new ClassLibrary1.Huangzilong.Run_huangzilong.rotor_assembly().rotor_355_assembly(rotor_assembly_code[0]);
            else if (housenumber == "280")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().转子280装配(rotor_assembly_code[0]);
            else if (housenumber == "200")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().转子200装配(rotor_assembly_code[0]);
            else if (housenumber == "180")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().转子180装配(rotor_assembly_code[0]);



            //MessageBox.Show("转子结束");


            string[] sl, ring, clasp, stator_core_assembly;

            sl = Module_wangbo.stator_lamination_mysql(localhost, alihost, alitable_name);
            if (sl[1] == "repeat" && Module_huangzilong.check_file(localhost, "stator_lamination", "all_domeslot_lamination", sl[0],"part_save_address")) ;
            else
            {
                new ClassLibrary1.Run_wangbo.Stator_lamination().all_domeslot_lamination(sl[0], localhost);
            }

            ring = Module_wangbo.stator_pressing_ring_mysql(localhost, alitable_name);
            if (ring[1] == "repeat" && Module_huangzilong.check_file(localhost, "stator_pressing_ring", "stator_pressing_ring", ring[0], "part_save_address")) ;
            else
            {
                new ClassLibrary1.Run_wangbo.LY_synchronous_motor.Stator().stator_pressing_ring(ring[0], localhost);
            }

            clasp = Module_wangbo.clasp_mysql(localhost,sl[0],ring[0]);
            if (clasp[1] == "repeat" && Module_huangzilong.check_file(localhost, "clasp", "clasp", clasp[0], "part_save_address")) ;
            else
            {
                new ClassLibrary1.Run_wangbo.LY_synchronous_motor.Stator().clasp(clasp[0], localhost);
            }

            stator_core_assembly = Module_wangbo.stator_core_assembly_mysql(localhost, alihost, alitable_name);
            if (stator_core_assembly[1] == "repeat" && Module_huangzilong.check_file(localhost, "stator_core_assembly", "stator_core_assembly", stator_core_assembly[0], "part_save_address")) ;
            else
            {
                new ClassLibrary1.Run_wangbo.LY_synchronous_motor.Stator().stator_core_assembly(stator_core_assembly[0], localhost, clasp[0]);
            }




            //励磁机定子装配
            string[] Exciter_stator_core_assembly;
            Exciter_stator_core_assembly = Module_huangzilong.Exciter_stator_core_assembly(housenumber);



            //前端盖
            string[] front_end_cover_code;
            front_end_cover_code = Module_huangzilong.front_end_cover(housenumber);



            string[] main_frame;
            main_frame = Module_huangzilong.main_frame(housenumber);
            //MessageBox.Show("定子结束");




            string General_assembly_address, General_pack_address, pdf_pack_address;
            General_assembly_address = "E:\\works\\generator_parts_library\\generator_assembly\\";
            General_pack_address = "C:\\Users\\Administrator\\Desktop\\pack\\";
            pdf_pack_address = "C:\\Users\\Administrator\\Desktop\\pdf\\";

            //总装配
            string[] generator_assembly_code;


            generator_assembly_code = new ClassLibrary1.Huangzilong.Module_huangzilong().generator_assembly(localhost, "generator_assembly", "generator_assembly", rotor_assembly_code[0], stator_core_assembly[0], Exciter_stator_core_assembly[0], front_end_cover_code[0], main_frame[0], housenumber);


            Module_huangzilong.修改筒体长度(Shaft_code[0], housenumber, localhost);
            Module_huangzilong.修改筋长度(Shaft_code[0], housenumber, localhost);
            if ("repeat" == generator_assembly_code[1] && new ClassLibrary1.Huangzilong.Module_huangzilong().check_file(localhost, "generator_assembly", "generator_assembly", generator_assembly_code[0])) ;

            else if (housenumber == "355")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配355(generator_assembly_code[0]);
            else if (housenumber == "250")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配250(generator_assembly_code[0]);
            else if (housenumber == "280")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配280(generator_assembly_code[0]);
            else if (housenumber == "200")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配200(generator_assembly_code[0]);
            else if (housenumber == "180")
                new ClassLibrary1.fubin.Run_fubing.Parts_drawing().总装配180(generator_assembly_code[0]);
            //MessageBox.Show("555");

            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_and_go(General_assembly_address + generator_assembly_code[0], General_pack_address);



            //定子冲片pdf
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + sl[0] + "_LY", pdf_pack_address + "LY_" + sl[0] + "_LY");
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + sl[0] + "_LY", pdf_pack_address + "LY_" + sl[0] + "_LY", 3);

            //前端盖pdf
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + front_end_cover_code[0] + "_LY", pdf_pack_address + "LY_" + front_end_cover_code[0] + "_LY");
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + front_end_cover_code[0] + "_LY", pdf_pack_address + "LY_" + front_end_cover_code[0] + "_LY", 3);

            //风扇pdf
            if (housenumber == "355" && housenumber == "280" && housenumber == "200")
            {
                new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + fan_code[0] + "_LY", pdf_pack_address + "LY_" + fan_code[0] + "_LY", 2);
                new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + fan_code[0] + "_LY", pdf_pack_address + "LY_" + fan_code[0] + "_LY", 3);
            }
            else
            {
                new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + fan_code[0] + "_LY", pdf_pack_address + "LY_" + fan_code[0] + "_LY", 1);
                new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + fan_code[0] + "_LY", pdf_pack_address + "LY_" + fan_code[0] + "_LY", 3);
            }

            //后端盖pdf
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + Exciter_stator_core_assembly[2] + "_LY", pdf_pack_address + "LY_" + Exciter_stator_core_assembly[2] + "_LY");
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + Exciter_stator_core_assembly[2] + "_LY", pdf_pack_address + "LY_" + Exciter_stator_core_assembly[2] + "_LY", 3);

            //转子冲片pdf
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + rotor_core_code[0] + "_LY", pdf_pack_address + "LY_" + rotor_core_code[0] + "_LY");
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + rotor_core_code[0] + "_LY", pdf_pack_address + "LY_" + rotor_core_code[0] + "_LY", 3);

            //转轴
            new ClassLibrary1.Huangzilong.Module_huangzilong().pack_to_pdf(General_pack_address + "LY_" + Shaft_code[0] + "_LY", pdf_pack_address + "LY_" + Shaft_code[0] + "_LY");
              
             

        }




    

        protected void 定制铁芯(object sender, EventArgs e)
        {
            //new ClassLibrary1.Huangzilong.Run_huangzilong.定制铁芯().定制铁芯("001");//铁芯
            //new ClassLibrary1.Huangzilong.Run_huangzilong.定制铁芯().阻尼板("002");//阻尼板
            new ClassLibrary1.Huangzilong.Run_huangzilong.Damping_rod().Damping_rod("8LY.549.002");//阻尼棒
            new ClassLibrary1.Huangzilong.Run_huangzilong.定制铁芯().rotor_core_assembly();//装配

        }

        protected void 同步发电机转子装配(object sender, EventArgs e)
        {

        }
    }
}