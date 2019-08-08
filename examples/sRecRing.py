import hycohanz as hfss

raw_input('Press "Enter" to connect to HFSS.>')

with hfss.App() as App:
    raw_input('Press "Enter" to create a new project.>')

    with hfss.NewProject(App.oDesktop) as P:
        raw_input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

        with hfss.InsertDesign(P.oProject, "HFSSDesign1", "DrivenModal") as D:
            raw_input('Press "Enter" to set the active editor to "3D Modeler" (The default and only known correct value).>')

            with hfss.SetActiveEditor(D.oDesign) as E:

                dL = 0.4
                thick = 1.6

                hfss.add_property_project(P.oProject, "$L", "10mm")
                hfss.add_property_project(P.oProject, "$Theta", "0deg")

                raw_input('Press "Enter" to draw a red box named Box1.>')

                box1 = hfss.create_box_new(E.oEditor,
                                '0mm',
                                '0mm',
                                '-15mm',
                                '$L+'+str(2*dL)+'mm',
                                '$L+'+str(2*dL)+'mm',
                                '50mm',
                                'mm',
                                0, 0, -15, 10+2*dL, 10+2*dL, 50,
                                Name='Box1',
                                Color=(255, 0, 0))

                raw_input('Press "Enter" to draw a red box named Box2.>')

                box2 = hfss.create_box_new(E.oEditor,
                                '0mm',
                                '0mm',
                                '0mm',
                                '$L+'+str(2*dL)+'mm',
                                '$L+'+str(2*dL)+'mm',
                                str(-thick)+'mm',
                                'mm',
                                0, 0, 0, 10+2*dL, 10+2*dL, -thick,
                                Name='Box2',
                                Color=(255, 0, 0),
                                MaterialValue='"FR4_epoxy"')

                '''raw_input('Press "Enter" to draw a red box named Box3.>')

                box3 = hfss.create_box(E.oEditor,
                                '0mm',
                                '0mm',
                                str(-thick)+'mm',
                                '$L+'+str(2*dL)+'mm',
                                '$L+'+str(2*dL)+'mm',
                                '-10mm',
                                'mm',
                                0, 0, -thick, 10+2*dL, 10+2*dL, -10,
                                Name='Box3',
                                Color=(255, 0, 0))'''

                raw_input('Press "Enter" to draw a red Rectangle named Rectangle1.>')

                r1 = hfss.create_rectangle(E.oEditor,
                                      str(dL)+'mm',
                                      str(dL)+'mm',
                                      '0mm',
                                      '$L',
                                      '$L',
                                      Name='Rectangle1',
                                      Color=(255, 0, 0))

                raw_input('Press "Enter" to draw a red Rectangle named Rectangle2.>')

                r2 = hfss.create_rectangle(E.oEditor,
                                           str(2*dL)+'mm',
                                           str(2*dL)+'mm',
                                           '0mm',
                                           '$L-'+str(2*dL)+'mm',
                                           '$L-'+str(2*dL)+'mm',
                                           Name='Rectangle2',
                                           Color=(255, 0, 0))

                raw_input('Press "Enter" to subtract the second Rectangle from the first.>')

                hfss.subtract(E.oEditor, [r1], [r2])

                raw_input('Press "Enter" to assign a PerfectE boundary condition on the ring and baseboard.>')

                ring_id = hfss.get_face_by_position(E.oEditor, r1, str(1.5*dL)+'mm', str(1.5*dL)+'mm', 0)
                # p = box2.get_face_point('down')
                # base_id = hfss.get_face_by_position(E.oEditor, box2.name, p[0], p[1], p[2])
                hfss.assign_perfect_e(D.oDesign, "PerfectE1", [ring_id])

                raw_input('Press "Enter" to assign a Master boundary condition on a face of Box1.>')

                p = box1.get_face_point('front')
                e = box1.get_face_edge('front')
                master_id_1 = hfss.get_face_by_position(E.oEditor, box1.name, p[0], p[1], p[2])
                hfss.assign_master(D.oDesign, "Master1", [master_id_1],
                                   e[0], e[1])

                p = box1.get_face_point('rear')
                e = box1.get_face_edge('rear')
                slave_id_1 = hfss.get_face_by_position(E.oEditor, box1.name, p[0],p[1],p[2])
                hfss.assign_slave(D.oDesign, "Master1", "Slave1", [slave_id_1],
                                  e[0],e[1])

                raw_input('Press "Enter" to assign a Floquet Port.>')

                p = box1.get_face_point('up')
                e = box1.get_face_edge('up')
                floquet_id1 = hfss.get_face_by_position(E.oEditor, box1.name, p[0],p[1],p[2])
                hfss.assign_floquet(D.oDesign, "Floquet1", [floquet_id1],
                                    e[0],e[1],e[2],e[3])

                p = box1.get_face_point('down')
                e = box1.get_face_edge('down')
                floquet_id2 = hfss.get_face_by_position(E.oEditor, box1.name, p[0], p[1], p[2])
                hfss.assign_floquet(D.oDesign, "Floquet2", [floquet_id2],
                                    e[0], e[1], e[2], e[3])

                p = box1.get_face_point('right')
                e = box1.get_face_edge('right')
                master_id_2 = hfss.get_face_by_position(E.oEditor, box1.name, p[0], p[1], p[2])
                hfss.assign_master(D.oDesign, "Master2", [master_id_2],
                                   e[0], e[1])

                p = box1.get_face_point('left')
                e = box1.get_face_edge('left')
                slave_id_2 = hfss.get_face_by_position(E.oEditor, box1.name, p[0], p[1], p[2])
                hfss.assign_slave(D.oDesign, "Master2", "Slave2", [slave_id_2],
                                  e[0], e[1], Phi="90deg", Theta="$Theta")

                raw_input('Press "Enter" to insert a setup.>')

                hfss.insert_analysis_setup(D.oDesign, 2, Name='Setup1')

                raw_input('Press "Enter" to insert a frequency sweep.>')

                hfss.insert_frequency_sweep_linear_interpolating(D.oDesign, 'Setup1', 'Sweep1', 1, 8, 0.5)

                raw_input('Press "Enter" to insert a optimetrics setup.>')

                hfss.insert_optimetrics_setup(D.oDesign, 'Setup1', ["L", "Theta"], ["LIN 6mm 16mm 0.05mm", 'LIN 0deg 60deg 10deg'], op_name="ParametricSetup1")

                raw_input('Press "Enter" to start simulating.>')

                hfss.solve_optimetrics(D.oDesign, 'ParametricSetup1')

                raw_input('Press "Enter" to create report.>')

                hfss.create_report(D.oDesign, ['Freq', '$L', '$Theta'], [['All'], ['All'], ['All']], "Freq",
                                   ["dB(S(Floquet1:1,Floquet2:1))"], "Setup1 : Sweep1", report_name="Plot1")
                hfss.export_to_file(D.oDesign, "Plot1", "C:\Users\Administrator\Desktop\Test.txt")

                file_name = raw_input('Please input file name:(Press Enter to save it as sRecRing.hfss to Desktop)')

                if file_name == '':
                    file_name = 'C:\Users\Administrator\Desktop\sRecRing_highaccuracy.hfss'

                hfss.save_as_project(App.oDesktop, file_name)

                raw_input('Press "Enter" to quit HFSS.>')

